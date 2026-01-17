let
    TRNTYPE_MAP_DATA =
        Table.NestedJoin(
            CLEANED_DATA,
            {"TransTypeStd"},
            TRNTYPE_MAP,
            {"RawTransType"},
            "Map",
            JoinKind.LeftOuter
        ),

    Expanded =
        Table.ExpandTableColumn(
            TRNTYPE_MAP_DATA,
            "Map",
            {"TxnClass","CashRule","UnitsRule","CostRule","IncomeType","ReturnType","RequiresSecurityKey","Notes"},
            {"TxnClass","CashRule","UnitsRule","CostRule","IncomeType","ReturnType","RequiresSecurityKey","TxnNotes"}
        ),

    // Explicit unmapped flag (exception will be built downstream)
    AddUnmapped =
        Table.AddColumn(
            Expanded,
            "IsUnmapped",
            each [TxnClass] = null,
            type logical,

    AddPeriodMonth = Table.AddColumn(AddUnmapped, "PeriodMonth", each Date.StartOfMonth([ValueDate]), type date),
    AddPeriodQuarter =
        Table.AddColumn(
            AddPeriodMonth,
            "PeriodQuarter",
            each #date(Date.Year([ValueDate]), Number.RoundUp(Date.Month([ValueDate]) / 3) * 3 - 2, 1),
            type date
        ),
    AddPeriodYear = Table.AddColumn(AddPeriodQuarter, "PeriodYear", each Date.Year([ValueDate]), Int64.Type),

    AddCashDelta =
        Table.AddColumn(
            AddPeriodYear,
            "CashDelta",
            each
                let ta = [TotalAmount],
                    cm = if [Commission] = null then 0 else [Commission]
                in
                    if [CashRule] = "TOTAL_PLUS_COMM" then (if ta = null then null else ta + cm)
                    else if [CashRule] = "TOTAL" then ta
                    else if [CashRule] = "NONE" then 0
                    else null,
            type number
        ),

    AddUnitsDelta =
        Table.AddColumn(
            AddCashDelta,
            "UnitsDelta",
            each if [UnitsRule] = "UNITS" then [Units] else null,
            type number
        ),

    AddCostDeltaPre =
        Table.AddColumn(
            AddUnitsDelta,
            "CostDelta_Pre",
            each
                if [CostRule] = "BUY_COST" then
                    // prefer cash-derived cost when cash exists; else fall back to UnitPrice if available
                    if [CashDelta] <> null and [CashDelta] <> 0 then -[CashDelta]
                    else if [UnitsDelta] <> null and [UnitPrice] <> null then [UnitsDelta] * [UnitPrice]
                    else null
                else if [CostRule] = "REINVEST_COST" then
                    if [UnitsDelta] <> null and [ReinvestPrice] <> null then [UnitsDelta] * [ReinvestPrice] else null
                else if [CostRule] = "ZERO" then 0
                else null,
            type number
        ),

    // --- period open/close markers at ACCOUNT x MONTH level ---
    GroupMinMax =
        Table.Group(
            AddCostDeltaPre,
            {"Account","PeriodMonth"},
            {
                {"MinSort", each List.Min([LedgerRowID]), Int64.Type},
                {"MaxSort", each List.Max([LedgerRowID]), Int64.Type}
            }
        ),

    MergeMinMax =
        Table.NestedJoin(
            AddCostDeltaPre,
            {"Account","PeriodMonth"},
            GroupMinMax,
            {"Account","PeriodMonth"},
            "MM",
            JoinKind.LeftOuter
        ),

    ExpandMinMax =
        Table.ExpandTableColumn(MergeMinMax, "MM", {"MinSort","MaxSort"}, {"MinSort","MaxSort"}),

    AddIsOpen =
        Table.AddColumn(
            ExpandMinMax,
            "IsPeriodOpenRow",
            each [LedgerRowID] = [MinSort],
            type logical
        ),

    AddIsClose =
        Table.AddColumn(
            AddIsOpen,
            "IsPeriodCloseRow",
            each [LedgerRowID] = [MaxSort],
            type logical
        ),

    DropMM = Table.RemoveColumns(AddIsClose, {"MinSort","MaxSort"})
in
    DropMM