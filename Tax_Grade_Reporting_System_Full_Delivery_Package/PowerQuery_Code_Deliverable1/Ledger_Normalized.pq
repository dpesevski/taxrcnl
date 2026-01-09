// Ledger_Normalized : normalize schema + parse dates
let
    Source = Ledger_Raw,
    Renamed = Table.RenameColumns(
        Source,
        {
            {"TRANS/VALUE DATE","TransValueDate_Text"},
            {"ENTRY DATE","EntryDate_Text"},
            {"TRANS TYPE","RawTransType"},
            {"A/C","Account"},
            {"ST","Status"},
            {"QUANTITY","Quantity"},
            {"TRADE NO","TradeNo"},
            {"CUSIP","CUSIP"},
            {"SYMBOL","Symbol"},
            {"SECURITY DESCRIPTION","SecurityDescription"},
            {"AVG COST","AvgCost"},
            {"UNIT PRICE","UnitPrice"},
            {"TOTAL AMOUNT","TotalAmount"},
            {"REMARK","Remark"},
            {"GAIN/LOSS","GainLoss_Broker"},
            {"CXLD","CancelledFlag"},
            {"LONG REMARK","LongRemark"},
            {"COMMISSION","Commission"}
        },
        MissingField.Ignore
    ),
    ParsedTV = Table.AddColumn(Renamed, "TV", each fnParseTransValueDate([TransValueDate_Text])),
    ExpandedTV = Table.ExpandRecordColumn(ParsedTV, "TV", {"TransactionDate","ValueDate","ParseOk","Raw"}, {"TransactionDate","ValueDate","TV_ParseOk","TransValueDate_Raw"}),
    ParsedEntry = Table.AddColumn(ExpandedTV, "EntryDate", each try Date.FromText([EntryDate_Text]) otherwise null, type date),
    // Coerce numerics (keep null if non-parsable; exceptions will capture)
    Num = Table.TransformColumns(
        ParsedEntry,
        {
            {"Quantity", each try Number.FromText(_) otherwise null, type number},
            {"TotalAmount", each try Number.FromText(_) otherwise null, type number},
            {"Commission", each try Number.FromText(_) otherwise null, type number},
            {"UnitPrice", each try Number.FromText(_) otherwise null, type number},
            {"AvgCost", each try Number.FromText(_) otherwise null, type number},
            {"GainLoss_Broker", each try Number.FromText(_) otherwise null, type number}
        },
        MissingField.Ignore
    ),
    SecurityKey = Table.AddColumn(Num, "SecurityKey", each if [CUSIP] <> null and Text.Trim(Text.From([CUSIP])) <> "" then Text.Trim(Text.From([CUSIP])) else Text.Trim(Text.From([Symbol])), type text),
    LedgerRowID = Table.AddColumn(SecurityKey, "LedgerRowID", each Text.From([SourceName]) & ":" & Text.From([RawRowNum]), type text)
in
    LedgerRowID