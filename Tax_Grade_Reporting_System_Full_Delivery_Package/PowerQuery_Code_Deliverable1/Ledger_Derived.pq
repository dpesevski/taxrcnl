// Ledger_Derived : apply mapping rules + derive CashDelta/UnitsDelta/CostDelta (pre-ACB)
let
    Source = Ledger_Normalized,
    Map = Excel.CurrentWorkbook(){[Name="tblTxnTypeMap"]}[Content],
    MapKeyed = Table.TransformColumnTypes(Map, {{"RawTransType", type text}, {"TxnClass", type text}}),
    Joined = Table.NestedJoin(Source, {"RawTransType"}, MapKeyed, {"RawTransType"}, "Map", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Map", {"TxnClass","IncomeType","ReturnType","TransferType"}, {"TxnClass","IncomeType","ReturnType","TransferType"}),
    TxnClassFilled = Table.ReplaceValue(Expanded, null, "UNMAPPED", Replacer.ReplaceValue, {"TxnClass"}),
    IncomeTypeFilled = Table.ReplaceValue(TxnClassFilled, null, "UNMAPPED", Replacer.ReplaceValue, {"IncomeType"}),
    ReturnTypeFilled = Table.ReplaceValue(IncomeTypeFilled, null, "UNMAPPED", Replacer.ReplaceValue, {"ReturnType"}),

    // CashDelta: sign conventions assume TotalAmount positive in broker export for sells/income; buys may appear negative or positive depending on file.
    CashDelta = Table.AddColumn(ReturnTypeFilled, "CashDelta",
        each
            if [TotalAmount] = null then null
            else
                let
                    amt = [TotalAmount],
                    fee = if [Commission] = null then 0 else [Commission],
                    cls = [TxnClass]
                in
                    if cls = "SEC_BUY" then -Number.Abs(amt) - fee
                    else if cls = "SEC_SELL" then Number.Abs(amt) - fee
                    else if cls = "DIV" or cls = "INT" then Number.Abs(amt)
                    else if Text.StartsWith(cls, "CANCEL") then -amt // simplistic; exceptions will verify netting
                    else if cls = "UNMAPPED" then null
                    else amt
        , type number),

    UnitsDelta = Table.AddColumn(CashDelta, "UnitsDelta",
        each
            if [Quantity] = null then 0
            else
                let cls=[TxnClass] in
                if cls="SEC_BUY" then Number.Abs([Quantity])
                else if cls="SEC_SELL" then -Number.Abs([Quantity])
                else if cls="XFER_IN" then Number.Abs([Quantity])
                else if cls="XFER_OUT" then -Number.Abs([Quantity])
                else if Text.StartsWith(cls,"CANCEL") then -[Quantity]
                else 0
        , type number),

    // CostDelta preliminary (buys add, ROC reduces, others 0; sells handled in ACB engine)
    CostDeltaPre = Table.AddColumn(UnitsDelta, "CostDelta_Pre",
        each
            let cls=[TxnClass] in
            if cls="SEC_BUY" and [CashDelta] <> null then Number.Abs([CashDelta]) // cash out increases ACB
            else if cls="ROC" and [TotalAmount] <> null then -Number.Abs([TotalAmount])
            else if cls="UNMAPPED" then null
            else 0
        , type number),

    PeriodMonth = Table.AddColumn(CostDeltaPre, "PeriodMonth", each if [ValueDate] = null then null else Date.ToText(Date.StartOfMonth([ValueDate]), "yyyy-MM"), type text),
    PeriodQuarter = Table.AddColumn(PeriodMonth, "PeriodQuarter", each if [ValueDate] = null then null else Text.From(Date.Year([ValueDate])) & "-Q" & Text.From(Date.QuarterOfYear([ValueDate])), type text),
    PeriodYear = Table.AddColumn(PeriodQuarter, "PeriodYear", each if [ValueDate] = null then null else Date.Year([ValueDate]), Int64.Type)
in
    PeriodYear