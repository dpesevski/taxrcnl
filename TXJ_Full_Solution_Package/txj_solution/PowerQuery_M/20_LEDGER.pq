let
    Raw = RAW_Transactions,
    # Keep only if expected columns exist
    KeepCols = {
        "TransValueDate","EntryDate","TransType","AccountRaw","Quantity","TradeNo","CUSIP","Symbol","SecurityDescription",
        "AvgCost","UnitPrice","TotalAmount","CostAmount","Currency_TotalAmount","Remark","GainLoss","LongRemark","Commission",
        "RawRowID","SourceFileName","SourceFilePath","SourceFileModified","SourceFileExt"
    },
    Existing = List.Intersect({KeepCols, Table.ColumnNames(Raw)}),
    Trimmed = if List.Count(Existing) = 0 then Raw else Table.SelectColumns(Raw, Existing),

    # Trim text columns
    TextCols = List.Intersect({{"TransType","AccountRaw","CUSIP","Symbol","SecurityDescription","Remark","LongRemark","TradeNo","SourceFileName"}, Table.ColumnNames(Trimmed)}),
    TrimText = Table.TransformColumns(Trimmed, List.Transform(TextCols, each {_, Text.Trim, type text})),

    # Parse dates: TransValueDate may be "dd-MMM-yy/dd-MMM-yy" or Excel date
    AddTradeDate = Table.AddColumn(TrimText, "TradeDate", each
        let
            tv = [TransValueDate],
            tvt = if tv = null then null else Text.From(tv),
            parts = if tvt = null then {} else Text.Split(tvt, "/"),
            t0 = if List.Count(parts) = 0 then null else parts{0},
            d = try Date.From(tv) otherwise try Date.FromText(t0, "en-US") otherwise null
        in
            d, type date),

    AddSettleDate = Table.AddColumn(AddTradeDate, "SettleDate", each
        let
            tv = [TransValueDate],
            tvt = if tv = null then null else Text.From(tv),
            parts = if tvt = null then {} else Text.Split(tvt, "/"),
            t1 = if List.Count(parts) > 1 then parts{1} else (if List.Count(parts)>0 then parts{0} else null),
            d = try Date.From(tv) otherwise try Date.FromText(t1, "en-US") otherwise null
        in
            d, type date),

    AddEntryDate = Table.AddColumn(AddSettleDate, "EntryDateParsed", each try Date.From([EntryDate]) otherwise try Date.FromText(Text.From([EntryDate]), "en-US") otherwise null, type date),

    # Parse numbers
    NumCols = List.Intersect({{"Quantity","AvgCost","UnitPrice","TotalAmount","CostAmount","GainLoss","Commission"}, Table.ColumnNames(AddEntryDate)}),
    ParseNums = Table.TransformColumns(AddEntryDate, List.Transform(NumCols, each {_, (x) => try Number.From(x) otherwise try Number.FromText(Text.From(x)) otherwise null, type number})),

    # UnitPrice from LongRemark if missing
    AddParsedPrice = Table.AddColumn(ParseNums, "ParsedUnitPrice", each fnParsePriceFromText([LongRemark]), type number),
    AddUnitPriceFinal = Table.AddColumn(AddParsedPrice, "UnitPriceFinal", each if [UnitPrice] = null or [UnitPrice] = 0 then [ParsedUnitPrice] else [UnitPrice], type number),

    # Currency (default CAD if missing)
    AddCurrency = Table.AddColumn(AddUnitPriceFinal, "Currency", each
        let c = try Text.Upper(Text.Trim(Text.From([Currency_TotalAmount]))) otherwise null
        in if c = null or c = "" then "CAD" else c, type text),

    # Standard identifiers
    AddAccountID = Table.AddColumn(AddCurrency, "AccountID", each if [AccountRaw] = null then null else Text.Upper(Text.Trim(Text.From([AccountRaw]))), type text),
    AddSecurityID = Table.AddColumn(AddAccountID, "SecurityID", each
        let
            c = try Text.Upper(Text.Trim(Text.From([CUSIP]))) otherwise null,
            s = try Text.Upper(Text.Trim(Text.From([Symbol]))) otherwise null,
            d = try Text.Trim(Text.From([SecurityDescription])) otherwise null
        in
            if c <> null and c <> "" then c else if s <> null and s <> "" then s else d, type text),

    # TxnCodeRaw = TRANS TYPE | REMARK (uppercase)
    AddTxnCodeRaw = Table.AddColumn(AddSecurityID, "TxnCodeRaw", each
        let
            t = if [TransType] = null then "" else Text.Upper(Text.Trim([TransType])),
            r = if [Remark] = null then "" else Text.Upper(Text.Trim(Text.From([Remark])))
        in
            t & "|" & r, type text),

    # IncomeAmount (used for reinvested income when TotalAmount is zero)
    AddIncomeAmt = Table.AddColumn(AddTxnCodeRaw, "IncomeAmount", each
        let
            ta = try Number.From([TotalAmount]) otherwise null,
            q = try Number.From([Quantity]) otherwise null,
            up = try Number.From([UnitPriceFinal]) otherwise null
        in
            if ta <> null and ta <> 0 then ta
            else if q <> null and up <> null then q * up
            else 0, type number),

    # Join MAP_TxnType
    MapTxn = Excel.CurrentWorkbook(){[Name="tbl_Map_TxnType"]}[Content],
    MapTxn2 = Table.TransformColumns(MapTxn, {{"TxnCodeRaw", each Text.Upper(Text.Trim(Text.From(_))), type text}}),
    JoinTxn = Table.NestedJoin(AddIncomeAmt, {"TxnCodeRaw"}, MapTxn2, {"TxnCodeRaw"}, "TxnMap", JoinKind.LeftOuter),
    ExpandTxn = Table.ExpandTableColumn(JoinTxn, "TxnMap", {"NormalizedTxnType","CashRule","UnitsRule","CostRule","TransferType"}, {"NormalizedTxnType","CashRule","UnitsRule","CostRule","TransferType"}),
    FillTxn = Table.ReplaceValue(ExpandTxn, null, "UNMAPPED", Replacer.ReplaceValue, {"NormalizedTxnType"}),

    # Join MAP_TaxCategory
    MapTax = Excel.CurrentWorkbook(){[Name="tbl_Map_TaxCategory"]}[Content],
    JoinTax = Table.NestedJoin(FillTxn, {"NormalizedTxnType"}, MapTax, {"NormalizedTxnType"}, "TaxMap", JoinKind.LeftOuter),
    ExpandTax = Table.ExpandTableColumn(JoinTax, "TaxMap", {"TaxCategory"}, {"TaxCategory"}),
    FillTax = Table.ReplaceValue(ExpandTax, null, "UNMAPPED", Replacer.ReplaceValue, {"TaxCategory"}),

    # Apply rules for deltas
    UCaseRules = Table.TransformColumns(FillTax, {{"CashRule", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text}, {"UnitsRule", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text}, {"CostRule", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text}}),

    AddCashDelta = Table.AddColumn(UCaseRules, "CashDelta", each
        let
            rule = [CashRule],
            ta = try Number.From([TotalAmount]) otherwise null,
            ia = try Number.From([IncomeAmount]) otherwise null
        in
            if rule = "TOTALAMOUNT" then ta
            else if rule = "NEG_TOTALAMOUNT" then (if ta=null then null else -ta)
            else if rule = "INCOMEAMOUNT" then ia
            else if rule = "0" then 0
            else null, type number),

    AddUnitsDelta = Table.AddColumn(AddCashDelta, "UnitsDelta", each
        let
            rule = [UnitsRule],
            q = try Number.From([Quantity]) otherwise null
        in
            if rule = "QUANTITY" then q
            else if rule = "ABSQUANTITY" then (if q=null then null else Number.Abs(q))
            else if rule = "0" then 0
            else null, type number),

    AddCostDeltaPre = Table.AddColumn(AddUnitsDelta, "CostDeltaPre", each
        let
            rule = [CostRule],
            cd = try Number.From([CashDelta]) otherwise null,
            ia = try Number.From([IncomeAmount]) otherwise null,
            comm = try Number.From([Commission]) otherwise 0
        in
            if rule = "BUYCOST" then (if cd=null then null else (-cd) + comm)
            else if rule = "INCOMEREINV" then ia
            else if rule = "ROC" then (if ia=null then null else -ia)
            else if rule = "0" then 0
            else null, type number),

    # Period keys
    AddPeriodMonth = Table.AddColumn(AddCostDeltaPre, "PeriodMonth", each if [SettleDate]=null then null else Date.ToText([SettleDate], "yyyy-MM"), type text),
    AddPeriodYear = Table.AddColumn(AddPeriodMonth, "PeriodYear", each if [SettleDate]=null then null else Date.Year([SettleDate]), Int64.Type)

in
    AddPeriodYear
