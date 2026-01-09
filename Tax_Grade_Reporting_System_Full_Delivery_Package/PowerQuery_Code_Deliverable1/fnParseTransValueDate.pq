// fnParseTransValueDate
(txt as nullable any) as record =>
let
    t0 = if txt = null then null else Text.Trim(Text.From(txt)),
    // Supports formats like "2025-01-02/2025-01-05" or a single date "2025-01-02"
    parts =
        if t0 = null or t0 = "" then {}
        else List.Transform(Text.Split(t0, "/"), each Text.Trim(_)),

    dTrans =
        if List.Count(parts) >= 1
        then (try Date.FromText(parts{0}) otherwise null)
        else null,

    dValue =
        if List.Count(parts) >= 2
        then (try Date.FromText(parts{1}) otherwise null)
        else dTrans,

    ok = (dTrans <> null and dValue <> null)
in
    [TransactionDate = dTrans, ValueDate = dValue, ParseOk = ok, Raw = t0]