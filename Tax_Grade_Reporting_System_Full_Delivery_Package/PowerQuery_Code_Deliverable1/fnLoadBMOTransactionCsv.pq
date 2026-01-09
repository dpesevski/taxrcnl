// fnLoadBMOTransactionCsv
(fileContent as binary, fileName as text) as table =>
let
    // BMO exports in this project are tab-delimited with header row
    Source = Csv.Document(
        fileContent,
        [
            Delimiter = "#(tab)",
            Columns = 19,
            Encoding = 65001,
            QuoteStyle = QuoteStyle.None
        ]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Trimmed = Table.TransformColumns(Promoted, List.Transform(Table.ColumnNames(Promoted), (c)=> {c, each if _ is text then Text.Trim(_) else _, type any})),
    AddedSource = Table.AddColumn(Trimmed, "SourceName", each fileName, type text),
    AddedRowNum = Table.AddIndexColumn(AddedSource, "RawRowNum", 1, 1, Int64.Type)
in
    AddedRowNum