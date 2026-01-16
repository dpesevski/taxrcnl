let
    L = LEDGER,
    Seed = Excel.CurrentWorkbook(){[Name="tbl_Opening_Seed"]}[Content],
    SeedKeyed = Table.TransformColumns(Seed, {{"AccountID", each Text.Upper(Text.Trim(Text.From(_))), type text}, {"SecurityID", each Text.Upper(Text.Trim(Text.From(_))), type text}, {"Currency", each Text.Upper(Text.Trim(Text.From(_))), type text}}),

    # Helper function: ACB roll-forward on a sorted table
    fnACB = (t as table, openingUnits as number, openingACB as number) as table =>
        let
            Recs = Table.ToRecords(t),
            Out = List.Accumulate(Recs, {}, (state, cur) =>
                let
                    prevUnits = if List.Count(state)=0 then openingUnits else state{List.Count(state)-1}[RunningUnits],
                    prevACB = if List.Count(state)=0 then openingACB else state{List.Count(state)-1}[RunningACB],

                    unitsDelta = try Number.From(cur[UnitsDelta]) otherwise 0,
                    cashDelta = try Number.From(cur[CashDelta]) otherwise 0,
                    costDeltaPre = try Number.From(cur[CostDeltaPre]) otherwise 0,
                    comm = try Number.From(cur[Commission]) otherwise 0,
                    ntype = if Record.HasFields(cur, "NormalizedTxnType") then Text.Upper(Text.From(cur[NormalizedTxnType])) else "",

                    isSell = (ntype = "SELL"),
                    sellUnits = if isSell then Number.Abs(unitsDelta) else 0,
                    avgCost = if prevUnits = 0 then null else prevACB / prevUnits,
                    costRemoved = if isSell and avgCost <> null then sellUnits * avgCost else 0,

                    costAdd = if isSell then 0 else costDeltaPre,

                    runningACB = prevACB + costAdd - costRemoved,
                    runningUnits = prevUnits + unitsDelta,

                    realizedGain = if isSell and avgCost <> null then cashDelta - costRemoved - comm else null,

                    rec1 = Record.AddField(cur, "AvgCostPerUnit", avgCost),
                    rec2 = Record.AddField(rec1, "CostRemoved", costRemoved),
                    rec3 = Record.AddField(rec2, "RunningUnits", runningUnits),
                    rec4 = Record.AddField(rec3, "RunningACB", runningACB),
                    rec5 = Record.AddField(rec4, "RealizedGain", realizedGain)
                in
                    state & {rec5}
            ),
            Tbl = Table.FromRecords(Out)
        in
            Tbl,

    # Group by Account+Security+Currency and apply ACB
    AddKeys = Table.AddColumn(L, "_Key", each [AccountID] & "|" & [SecurityID] & "|" & [Currency], type text),

    Grouped = Table.Group(AddKeys, {"AccountID","SecurityID","Currency"}, {
        {"Data", (t) =>
            let
                sorted = Table.Sort(t, {{"SettleDate", Order.Ascending}, {"RawRowID", Order.Ascending}}),
                acc = t{0}[AccountID],
                sec = t{0}[SecurityID],
                cur = t{0}[Currency],
                seedRows = Table.SelectRows(SeedKeyed, each [AccountID]=acc and [SecurityID]=sec and [Currency]=cur),
                ou = if Table.RowCount(seedRows)=0 then 0 else try Number.From(seedRows{0}[OpeningUnits]) otherwise 0,
                oa = if Table.RowCount(seedRows)=0 then 0 else try Number.From(seedRows{0}[OpeningACB]) otherwise 0
            in
                fnACB(sorted, ou, oa)
        }
    }),

    Expanded = Table.ExpandTableColumn(Grouped, "Data", Table.ColumnNames(Grouped[Data]{0}), Table.ColumnNames(Grouped[Data]{0}))

in
    Expanded
