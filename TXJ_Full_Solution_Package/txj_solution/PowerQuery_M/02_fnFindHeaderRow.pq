let
    fnFindHeaderRow = (tbl as table, headerKey as text) as number =>
let
    key = Text.Upper(headerKey),
    rows = Table.ToRows(tbl),
    flags = List.Transform(rows, (r) =>
        let
            texts = List.Transform(r, each Text.Upper(Text.From(_)))
        in
            List.Contains(texts, key)
    ),
    idx = List.PositionOf(flags, true)
in
    if idx = -1 then 0 else idx

in
    fnFindHeaderRow
