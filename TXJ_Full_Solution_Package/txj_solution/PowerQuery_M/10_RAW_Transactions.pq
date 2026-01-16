let
    InputFolder = pInputFolder,
    Files = try Folder.Files(InputFolder) otherwise #table({}, {}),
    TxnFiles = if Table.IsEmpty(Files) then Files else Table.SelectRows(Files, each List.Contains({".csv", ".xlsx", ".xls"}, Text.Lower([Extension]))),
    WithData = if Table.IsEmpty(TxnFiles) then #table({}, {}) else Table.AddColumn(TxnFiles, "Data", each fnParseTransactionFile([Folder Path] & [Name], [Name], [Extension], [Date modified])),
    Expanded = if Table.IsEmpty(WithData) then #table({}, {}) else
        let
            sample = WithData[Data]{0},
            cols = Table.ColumnNames(sample)
        in
            Table.ExpandTableColumn(WithData, "Data", cols, cols)

in
    Expanded
