// EXCEPTIONS : fail-loud exception table
let
    L = ACB_ENGINE,
    // Basic exception flags (extend per Appendix A)
    Add1 = Table.AddColumn(L, "Ex_Unmapped", each [TxnClass] = "UNMAPPED", type logical),
    Add2 = Table.AddColumn(Add1, "Ex_BadDates", each [TV_ParseOk] <> true, type logical),
    Add3 = Table.AddColumn(Add2, "Ex_MissingAmount", each ([CashDelta]=null and ([TxnClass]="SEC_BUY" or [TxnClass]="SEC_SELL" or [IncomeType] <> "NON_INCOME")), type logical),
    Add4 = Table.AddColumn(Add3, "Ex_MissingACBForSale", each ([TxnClass]="SEC_SELL" and ([CostRemoved]=null or [CostRemoved]=0) and [UnitsDelta]<>0), type logical),
    Add5 = Table.AddColumn(Add4, "Ex_ROC_Missing", each ([TxnClass]="ROC" and [TotalAmount]=null), type logical),

    OnlyExceptions = Table.SelectRows(Add5, each [Ex_Unmapped] or [Ex_BadDates] or [Ex_MissingAmount] or [Ex_MissingACBForSale] or [Ex_ROC_Missing]),
    // Shape into exception log schema
    AddId = Table.AddIndexColumn(OnlyExceptions, "ExceptionID", 1, 1, Int64.Type),
    AddSeverity = Table.AddColumn(AddId, "Severity",
        each
            if [Ex_Unmapped] or [Ex_MissingACBForSale] then "Level 1 — Critical"
            else if [Ex_BadDates] or [Ex_ROC_Missing] then "Level 2 — Warning"
            else "Level 3 — Info"
    , type text),
    AddType = Table.AddColumn(AddSeverity, "ExceptionType",
        each
            if [Ex_Unmapped] then "UNMAPPED_TRANSACTION_TYPE"
            else if [Ex_MissingACBForSale] then "MISSING_ACB_FOR_SALE"
            else if [Ex_BadDates] then "BAD_DATE_PARSE"
            else if [Ex_ROC_Missing] then "ROC_MISSING_COMPONENT"
            else "OTHER"
    , type text),
    AddDesc = Table.AddColumn(AddType, "Description",
        each
            "TxnClass=" & Text.From([TxnClass]) & "; RawTransType=" & Text.From([RawTransType]) &
            "; LedgerRowID=" & Text.From([LedgerRowID])
    , type text),
    Selected = Table.SelectColumns(
        AddDesc,
        {"ExceptionID","Severity","Account","PeriodMonth","LedgerRowID","SourceName","RawRowNum","ExceptionType","Description"}
    )
in
    Selected