// WeatherLookup
let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMjIzMFDSUTJTitUBcwyBHFMwxzDN2MgcyDOG8izNTYE8IzjPDMgzhGpLBkmZKMXGAgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [HexValue = _t, SunIndex = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"HexValue", type text}, {"SunIndex", Int64.Type}})
in
    #"Changed Type"