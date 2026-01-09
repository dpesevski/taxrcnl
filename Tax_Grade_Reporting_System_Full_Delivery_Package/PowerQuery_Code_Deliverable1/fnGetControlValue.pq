// fnGetControlValue
(paramName as text) as any =>
let
    t = Excel.CurrentWorkbook(){[Name="tblControlLog"]}[Content],
    r = Table.SelectRows(t, each [Parameter] = paramName),
    v = if Table.RowCount(r) = 0 then null else r{0}[Value]
in
    v