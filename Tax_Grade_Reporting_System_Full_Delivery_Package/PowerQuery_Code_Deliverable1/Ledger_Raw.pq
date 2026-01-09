// Ledger_Raw : load all CSVs from FolderPath and combine
let
    FolderPath = fnGetControlValue("FolderPath"),
    Files = Folder.Files(FolderPath),
    Csvs = Table.SelectRows(Files, each Text.EndsWith([Extension], ".csv")),
    AddedTable = Table.AddColumn(Csvs, "Data", each fnLoadBMOTransactionCsv([Content],[Name])),
    Expanded = Table.ExpandTableColumn(AddedTable, "Data", Table.ColumnNames(AddedTable{0}[Data])),
    // Keep file metadata for provenance
    WithFileId = Table.AddColumn(Expanded, "FileId", each [Name] & " | " & DateTime.ToText([Date modified], "yyyy-MM-ddTHH:mm:ss"), type text)
in
    WithFileId