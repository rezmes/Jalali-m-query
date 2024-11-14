let
    Source =
        let
            func = (
                Source as table,
                StartDateColumn as text,
                optional EndDateColumn as text,
                optional FirstDayOfWeek as nullable number,
                optional StartEndAccuracy as any,
                optional Culture as any
            ) =>
                let
                    // Setup defaults for missing arguments
                    CLT = if Culture = null then "fa" else Culture,
                    EDC = if EndDateColumn = null then StartDateColumn else EndDateColumn,
                    FDW = if FirstDayOfWeek = null then 6 else FirstDayOfWeek,
                    // Saturday as the first day of the week
                    SEA = Text.Upper(Text.Range(Text.TrimStart(StartEndAccuracy), 0, 1)),
                    // Function to convert Gregorian date to Persian date
                    fnGregorianToPersian = (date as date) as record =>
                        let
                            PersianDate = DateTime.ToText(DateTime.From(date), "yyyy/MM/dd", "fa-IR"),
                            Parts = Text.Split(PersianDate, "/"),
                            Year = Number.FromText(Parts{0}),
                            Month = Number.FromText(Parts{1}),
                            Day = Number.FromText(Parts{2})
                        in
                            [Year = Year, Month = Month, Day = Day],
                    // Function to determine if a Persian year is a leap year
                    fnIsPersianLeapYear = (year as number) as logical =>
                        let
                            a = year - 474, b = 474 + Number.Mod(a, 2820), isLeap = Number.Mod(((b + 38) * 682), 2816) < 682
                        in
                            isLeap,
                    // Calculate the start and end date
                    StartCalendar = Number.RoundTowardZero(
                        (
                            Number.From(
                                if SEA = "Y" or SEA = "J" then
                                    Date.StartOfYear(List.Min(Table.Column(Source, StartDateColumn)))
                                else if SEA = "Q" then
                                    Date.StartOfQuarter(List.Min(Table.Column(Source, StartDateColumn)))
                                else if SEA = "M" then
                                    Date.StartOfMonth(List.Min(Table.Column(Source, StartDateColumn)))
                                else if SEA = "W" then
                                    Date.StartOfWeek(List.Min(Table.Column(Source, StartDateColumn)), Day.Saturday)
                                else
                                    List.Min(Table.Column(Source, StartDateColumn))
                            )
                        ),
                        0
                    ),
                    EndCalendar = Number.RoundTowardZero(
                        (
                            Number.From(
                                if SEA = "Y" or SEA = "J" then
                                    Date.EndOfYear(List.Max(Table.Column(Source, EDC)))
                                else if SEA = "Q" then
                                    Date.EndOfQuarter(List.Max(Table.Column(Source, EDC)))
                                else if SEA = "M" then
                                    Date.EndOfMonth(List.Max(Table.Column(Source, EDC)))
                                else if SEA = "W" then
                                    Date.EndOfWeek(List.Max(Table.Column(Source, EDC)), Day.Saturday)
                                else
                                    List.Max(Table.Column(Source, EDC))
                            )
                        ),
                        0
                    ),
                    // Create the calendar dates
                    Calendar = Table.TransformColumnTypes(
                        Table.FromList({StartCalendar..EndCalendar}, Splitter.SplitByNothing()),
                        {{"Column1", type date}}
                    ),
                    // Add the necessary columns
                    PersianDate = Table.AddColumn(Calendar, "PersianDate", each fnGregorianToPersian([Column1])),
                    Year = Table.AddColumn(PersianDate, "cnYear", each[PersianDate][Year], Int64.Type),
                    Quarter = Table.AddColumn(
                        Year, "cnQuarter", each "Q" & Text.From(Number.RoundUp([PersianDate][Month] / 3)), type text
                    ),
                    YearQuarter = Table.AddColumn(
                        Quarter, "cnYearQuarter", each Text.From([cnYear]) & "-" & [cnQuarter], type text
                    ),
                    Month = Table.AddColumn(YearQuarter, "cnMonth", each[PersianDate][Month], Int64.Type),
                    MonthName = Table.AddColumn(Month, "cnMonthname", each Date.MonthName([Column1], CLT), type text),
                    DaysInMonth = Table.AddColumn(
                        MonthName,
                        "cnDaysInMonth",
                        each
                            if[PersianDate][Month] <= 6 then
                                31
                            else if[PersianDate][Month] <= 11 then
                                30
                            else if fnIsPersianLeapYear([PersianDate][Year]) then
                                30
                            else
                                29,
                        Int64.Type
                    ),
                    YearMonth = Table.AddColumn(
                        DaysInMonth,
                        "cnYearMonth",
                        each Text.From([cnYear]) & "-" & Text.PadStart(Text.From([cnMonth]), 2, "0"),
                        type text
                    ),
                    WeekOfYear = Table.AddColumn(
                        YearMonth, "cnWeekOfYear", each Date.WeekOfYear([Column1], Day.Saturday), Int64.Type
                    ),
                    YearWeek = Table.AddColumn(
                        WeekOfYear,
                        "cnYearWeek",
                        each Text.From([cnYear]) & "-" & Text.PadStart(Text.From([cnWeekOfYear]), 2, "0"),
                        type text
                    ),
                    Day = Table.AddColumn(YearWeek, "cnDay", each[PersianDate][Day], Int64.Type),
                    DayName = Table.AddColumn(Day, "cnDayName", each Date.DayOfWeekName([Column1], CLT), type text),
                    DayOfWeek = Table.AddColumn(
                        DayName, "cnDayOfWeek", each Date.DayOfWeek([Column1], Day.Saturday) + 1, Int64.Type
                    ),
                    DayOfYear = Table.AddColumn(DayOfWeek, "cnDayOfYear", each Date.DayOfYear([Column1]), Int64.Type),
                    RenameDateCol = Table.RenameColumns(DayOfYear, {{"Column1", StartDateColumn}}),
                    // Rename the columns based on the culture
                    Language = Text.Lower(Text.Start(CLT, 2)),
                    tData = Table.FromRows(
                        {
                            {
                                "en",
                                "Year",
                                "Quarter",
                                "Year-Quarter",
                                "Month",
                                "Monthname",
                                "Days in Month",
                                "Year-Month",
                                "Week of Year",
                                "Year-Week",
                                "ISO Week",
                                "Year-ISO Week",
                                "Day",
                                "Dayname",
                                "Day of Week",
                                "Day of Year"
                            },
                            {
                                "fa",
                                "سال",
                                "ربع",
                                "سال-ربع",
                                "ماه",
                                "نام ماه",
                                "روزهای ماه",
                                "سال-ماه",
                                "هفته سال",
                                "سال-هفته",
                                "هفته ISO",
                                "سال-هفته ISO",
                                "روز",
                                "نام روز",
                                "روز هفته",
                                "روز سال"
                            }
                        },
                        {
                            "Language",
                            "cnYear",
                            "cnQuarter",
                            "cnYearQuarter",
                            "cnMonth",
                            "cnMonthname",
                            "cnDaysInMonth",
                            "cnYearMonth",
                            "cnWeekOfYear",
                            "cnYearWeek",
                            "cnIsoWeek",
                            "cnYearIsoWeek",
                            "cnDay",
                            "cnDayName",
                            "cnDayOfWeek",
                            "cnDayOfYear"
                        }
                    ),
                    tHeader = Table.SelectRows(tData, each ([Language] = Language)),
                    tCount =
                        if Table.RowCount(tHeader) = 0 then
                            Table.SelectRows(tData, each ([Language] = "en"))
                        else
                            Table.SelectRows(tData, each ([Language] = Language)),
                    tMapping = Table.Transpose(Table.DemoteHeaders(tCount)),
                    Renamed = Table.TransformColumnNames(
                        RenameDateCol,
                        each
                            List.Accumulate(
                                Table.ToRecords(tMapping),
                                _,
                                (state, current) => if state = current[Column1] then current[Column2] else state
                            )
                    )
                in
                    Renamed,
            documentation = [
                Documentation.Name = "fnCalendar",
                Documentation.Description = "Creates a calendar from a table",
                Documentation.LongDescription = Documentation.Description,
                Documentation.Category = "Calendar",
                Documentation.Source = "",
                Documentation.Version = "1.5.1",
                Documentation.Author = "rezmes",
                Documentation.Examples = {
                    [
                        Description = "
fnCalendar(Table1, ""Date"", null, null, ""Year"", ""fa"")
    ",
                        Code = "
Source          : The source query
StartDateColumn : Name of the column that contains the start dates
EndDateColumn   : Name of the column that contains the end dates (StartDateColumn if omitted)
FirstDayOfWeek  : omitted=Monday ISO Week (default), otherwise WeekOfYear 0=Sunday, 1=Monday, 2=Tuesday and so on
StartEndAccuracy: ""Jahr"", ""Quartal"", ""Monat"", ""Woche"" oder ""Tag"" (Standard)
StartEndAccuracy: ""Year"", ""Quarter"", ""Month"", ""Week"" or ""Day"" (default)
Culture         : ""en"" (default) or ""en-US"" or ""fa"" or ""fa-IR"" or any other locale
    ",
                        Result = ""
                    ]
                }
            ]
        in
            Value.ReplaceType(func, Value.ReplaceMetadata(Value.Type(func), documentation))
in
    Source
