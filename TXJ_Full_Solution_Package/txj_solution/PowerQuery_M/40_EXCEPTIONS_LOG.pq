let
    L = LEDGER,

    # 1) Unmapped transaction types
    Unmapped = Table.SelectRows(L, each [NormalizedTxnType] = "UNMAPPED" or [TaxCategory] = "UNMAPPED"),
    Unmapped2 = Table.AddColumn(Unmapped, "ExceptionCategory", each "UNMAPPED_TXN"),
    Unmapped3 = Table.AddColumn(Unmapped2, "Severity", each "L1-CRITICAL"),
    Unmapped4 = Table.AddColumn(Unmapped3, "Details", each "TxnCodeRaw not mapped in MAP_TxnType and/or MAP_TaxCategory."),

    # 2) Reinvested income with missing unit price (Stock Div / Interest where TotalAmount=0)
    NeedsPrice = Table.SelectRows(L, each ([TransType] <> null and Text.Upper(Text.From([TransType])) = "STOCK DIV" or Text.Upper(Text.From([TransType])) = "INTEREST") and ([TotalAmount] = 0 or [TotalAmount] = null) and ([UnitPriceFinal] = null or [UnitPriceFinal] = 0)),
    NeedsPrice2 = Table.AddColumn(NeedsPrice, "ExceptionCategory", each "MISSING_UNIT_PRICE"),
    NeedsPrice3 = Table.AddColumn(NeedsPrice2, "Severity", each "L1-CRITICAL"),
    NeedsPrice4 = Table.AddColumn(NeedsPrice3, "Details", each "Reinvested income requires explicit price in LONG REMARK to derive investment cost; otherwise CostDelta=0 and must be disclosed."),

    Combined = Table.Combine({
        Table.SelectColumns(Unmapped4, {"Severity","ExceptionCategory","AccountID","SecurityID","SettleDate","TxnCodeRaw","SourceFileName","RawRowID","Details"}),
        Table.SelectColumns(NeedsPrice4, {"Severity","ExceptionCategory","AccountID","SecurityID","SettleDate","TxnCodeRaw","SourceFileName","RawRowID","Details"})
    }),

    AddID = Table.AddIndexColumn(Combined, "ExceptionID", 1, 1, Int64.Type)

in
    AddID
