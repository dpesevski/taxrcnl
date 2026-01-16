let
    fnParsePriceFromText = (t as nullable text) as nullable number =>
let
    txt = if t = null then null else Text.From(t),
    pos = if txt = null then -1 else Text.PositionOf(txt, "$", Occurrence.Last),
    sub = if pos = -1 then null else Text.Range(txt, pos + 1),
    digits = if sub = null then null else Text.Select(sub, {"0".."9", ".", ","}),
    digits2 = if digits = null then null else Text.Replace(digits, ",", ""),
    val = try Number.FromText(digits2) otherwise null
in
    val

in
    fnParsePriceFromText
