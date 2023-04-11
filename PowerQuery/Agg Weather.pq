// Agg Weather
let
    Source = Excel.Workbook(File.Contents("C:\Users\sezzz\OneDrive\Documents\BloomWatch\Data\Allen Heritage Village Bloom Watch.xlsx"), null, true),
    Table1_Table = Source{[Item="Table1",Kind="Table"]}[Data],
    #"Changed Type0" = Table.TransformColumnTypes(Table1_Table,{{"ID", Int64.Type}, {"Start time", type datetime}, {"Completion time", type datetime}, {"Email", type text}, {"Name", type any}, {"Name2", type text}, {"Date", type date}, {"Time", type text}, {"Weather", type text}, {"Catmint", type text}, {"Chocolate Daisy", type text}, {"Green Santolina ", type text}, {"Sage", type text}, {"Pink Texas Skullcap", type text}, {"Rock Rose", type text}, {"Four Nerve Daisy", type text}, {"Fall Aster", type text}, {"Autumn Sage", type text}, {"Gregg's Mistflower", type text}, {"Prairie Verbena", type text}, {"Yarrow", type text}, {"Shasta Daisy", type text}, {"Turk's Cap", type text}, {"Lantana", type text}, {"Additional Comments2", type text}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Changed Type0",{"Weather", "ID"}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Removed Other Columns", "Weather", Splitter.SplitTextByDelimiter(";", QuoteStyle.Csv), {"Weather.1", "Weather.2", "Weather.3", "Weather.4"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Weather.1", type text}, {"Weather.2", type text}, {"Weather.3", type text}, {"Weather.4", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Weather.4"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Columns","",null,Replacer.ReplaceValue,{"Weather.3"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","",null,Replacer.ReplaceValue,{"Weather.2"}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Replaced Value1", {"ID"}, "Attribute", "Value"),
    #"Removed Columns1" = Table.RemoveColumns(#"Unpivoted Columns",{"Attribute"}),
    #"Added Conditional Column" = Table.AddColumn(#"Removed Columns1", "Custom", each if [Value] = "☀️" then 4 else 1),
    #"Added Custom" = Table.AddColumn(#"Added Conditional Column", "UnicodeValue", each Number.ToText(Character.ToNumber([Value]),"x")),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each true),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"UnicodeValue"}, WeatherLookup, {"HexValue"}, "Table", JoinKind.LeftOuter),
    #"Expanded Table" = Table.ExpandTableColumn(#"Merged Queries", "Table", {"SunIndex"}, {"Table.SunIndex"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Expanded Table",{"Value", "Custom", "UnicodeValue"}),
    #"Grouped Rows" = Table.Group(#"Removed Columns2", {"ID"}, {{"SunIndex", each List.Average([Table.SunIndex]), type nullable number}, {"countoficons", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows"

// WeatherLookup
let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMjIzMFDSUTJTitUBcwyBHFMwxzDN2MgcyDOG8izNTYE8IzjPDMgzhGpLBkmZKMXGAgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [HexValue = _t, SunIndex = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"HexValue", type text}, {"SunIndex", Int64.Type}})
in
    #"Changed Type"