// ACB_ENGINE : average-cost roll forward per Account + SecurityKey
let
    Source = Ledger_Derived,
    Sorted = Table.Sort(Source, {{"Account", Order.Ascending}, {"SecurityKey", Order.Ascending}, {"ValueDate", Order.Ascending}, {"RawRowNum", Order.Ascending}}),

    Grouped = Table.Group(
        Sorted,
        {"Account","SecurityKey"},
        {
            {"Rows", (t)=>
                let
                    recs = Table.ToRecords(t),
                    step = List.Accumulate(
                        recs,
                        [RunningUnits=0.0, RunningACB=0.0, Out={}],
                        (state, r) =>
                            let
                                cls = Record.FieldOrDefault(r,"TxnClass","UNMAPPED"),
                                unitsDelta = Record.FieldOrDefault(r,"UnitsDelta",0.0),
                                costPre = Record.FieldOrDefault(r,"CostDelta_Pre",0.0),
                                proceeds = if cls="SEC_SELL" then Number.Abs(Record.FieldOrDefault(r,"CashDelta",0.0)) else 0.0,
                                // cost removed on sale based on prior average cost
                                avgCostPerUnit = if state[RunningUnits]=0 then null else state[RunningACB]/state[RunningUnits],
                                costRemoved = if cls="SEC_SELL" and avgCostPerUnit <> null then Number.Abs(unitsDelta) * avgCostPerUnit else 0.0,
                                realizedGL = if cls="SEC_SELL" and avgCostPerUnit <> null then proceeds - costRemoved else 0.0,
                                newUnits = state[RunningUnits] + unitsDelta,
                                newACB =
                                    if cls="SEC_SELL" then state[RunningACB] - costRemoved
                                    else state[RunningACB] + costPre,
                                outRec = Record.AddField(
                                    Record.AddField(
                                        Record.AddField(
                                            Record.AddField(r, "RunningUnits", newUnits),
                                            "RunningACB", newACB
                                        ),
                                        "AvgCostPerUnit", if newUnits=0 then null else newACB/newUnits
                                    ),
                                    "CostRemoved", costRemoved
                                ),
                                outRec2 = Record.AddField(outRec, "RealizedGainLoss", realizedGL),
                                outList = state[Out] & {outRec2}
                            in
                                [RunningUnits=newUnits, RunningACB=newACB, Out=outList]
                    ),
                    outTable = Table.FromRecords(step[Out])
                in
                    outTable,
                type table
            }
        }
    ),

    Combined = Table.Combine(Grouped[Rows]),
    // Final CostDelta (pre for non-sales, negative on sale as costRemoved)
    CostDelta = Table.AddColumn(Combined, "CostDelta",
        each
            if [TxnClass] = "SEC_SELL" then -[CostRemoved]
            else [CostDelta_Pre]
        , type number)
in
    CostDelta