let
    L = LEDGER,
    Seed = Excel.CurrentWorkbook(){[Name="tbl_Opening_Seed"]}[Content],
    Control = Excel.CurrentWorkbook(){[Name="tbl_Control_Log"]}[Content],
    TolRow = try Control{[Key="FXToleranceCAD"]} otherwise null,
    Tol = if TolRow = null then 0.50 else try Number.From(TolRow[Value]) otherwise 0.50,

    # Cash reconciliation (by account + currency)
    CashAgg = Table.Group(L, {"AccountID","Currency"}, {{"NetCashDelta", each List.Sum(List.RemoveNulls([CashDelta])), type number}}),
    SeedCash = Table.SelectRows(Seed, each [Type] = "CASH"),
    SeedCash2 = Table.TransformColumns(SeedCash, {{"AccountID", each Text.Upper(Text.Trim(Text.From(_))), type text}, {"Currency", each Text.Upper(Text.Trim(Text.From(_))), type text}}),

    Join = Table.NestedJoin(CashAgg, {"AccountID","Currency"}, SeedCash2, {"AccountID","Currency"}, "Seed", JoinKind.LeftOuter),
    Expand = Table.ExpandTableColumn(Join, "Seed", {"OpeningCash"}, {"OpeningCash"}),
    FillOpen = Table.ReplaceValue(Expand, null, 0, Replacer.ReplaceValue, {"OpeningCash"}),
    AddComputedClose = Table.AddColumn(FillOpen, "ComputedClosingCash", each [OpeningCash] + [NetCashDelta], type number),

    # In this template we may not have a holdings-based closing cash; expose computed close and require statement tie-out
    AddStatus = Table.AddColumn(AddComputedClose, "ReconStatus", each "REVIEW_REQUIRED"),
    AddNotes = Table.AddColumn(AddStatus, "ReconNotes", each "Provide statement closing cash in RAW_Holdings (or an EndingCash table) to mechanically validate."),

    Out = AddNotes
in
    Out
