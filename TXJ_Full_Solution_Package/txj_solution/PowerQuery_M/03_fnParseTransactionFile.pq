let
    fnParseTransactionFile = (filePath as text, fileName as text, extension as text, modified as datetime) as table =>
let
    ext = Text.Lower(extension),
    raw = if ext = ".csv" then
            Csv.Document(File.Contents(filePath), [Delimiter=",", Encoding=1252, QuoteStyle=QuoteStyle.Csv])
          else
            let
                wb = Excel.Workbook(File.Contents(filePath), null, true),
                # Prefer first visible sheet
                sheets = Table.SelectRows(wb, each [Kind] = "Sheet"),
                data = if Table.RowCount(sheets) > 0 then sheets{0}[Data] else wb{0}[Data]
            in
                data,

    headerRow = fnFindHeaderRow(raw, "TRANS/VALUE DATE"),
    skipped = Table.Skip(raw, headerRow),
    promoted = Table.PromoteHeaders(skipped, [PromoteAllScalars=true]),
    colTrim = Table.TransformColumnNames(promoted, each Text.Trim(_)),

    # Standardize column names if present
    renamePairs = {
        {"TRANS/VALUE DATE", "TransValueDate"},
        {"ENTRY DATE", "EntryDate"},
        {"TRANS TYPE", "TransType"},
        {"A/C", "AccountRaw"},
        {"QUANTITY", "Quantity"},
        {"TRADE NO", "TradeNo"},
        {"CUSIP", "CUSIP"},
        {"SYMBOL", "Symbol"},
        {"SECURITY DESCRIPTION", "SecurityDescription"},
        {"AVG COST", "AvgCost"},
        {"UNIT PRICE", "UnitPrice"},
        {"TOTAL AMOUNT", "TotalAmount"},
        {"REMARK", "Remark"},
        {"GAIN/LOSS", "GainLoss"},
        {"LONG REMARK", "LongRemark"},
        {"COMMISSION", "Commission"},
        {"Cost amount", "CostAmount"},
        {"Currency (total amt)", "Currency_TotalAmount"},
        {"Currency (gain/loss)", "Currency_GainLoss"},
        {"Currency (avg cost)", "Currency_AvgCost"}
    },

    renamed = Table.RenameColumns(colTrim, renamePairs, MissingField.Ignore),

    # Add file metadata + RawRowID
    idx = Table.AddIndexColumn(renamed, "RawRowID", 1, 1, Int64.Type),
    addName = Table.AddColumn(idx, "SourceFileName", each fileName, type text),
    addPath = Table.AddColumn(addName, "SourceFilePath", each filePath, type text),
    addMod = Table.AddColumn(addPath, "SourceFileModified", each modified, type datetime),
    addExt = Table.AddColumn(addMod, "SourceFileExt", each ext, type text)

in
    addExt

in
    fnParseTransactionFile
