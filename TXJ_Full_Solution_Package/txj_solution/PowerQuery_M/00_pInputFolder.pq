let
    Control = Excel.CurrentWorkbook(){[Name="tbl_Control_Log"]}[Content],
    Row = try Control{[Key="InputFolder"]} otherwise null,
    FolderPath = if Row = null then "C:\\TaxReporting\\Input\\" else Text.From(Row[Value])
in
    FolderPath
