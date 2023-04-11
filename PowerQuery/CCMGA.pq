// CCMGA
let
    Source = Excel.Workbook(File.Contents("C:\Users\sezzz\OneDrive\Documents\BloomWatch\Data\Allen Heritage Village Bloom Watch.xlsx"), null, true),
    Table1_Table = Source{[Item="Table1",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table1_Table,{{"ID", Int64.Type}, {"Start time", type datetime}, {"Completion time", type datetime}, {"Email", type text}, {"Name", type any}, {"Name2", type text}, {"Date", type date}, {"Time", type text}, {"Weather", type text}, {"Catmint", type text}, {"Chocolate Daisy", type text}, {"Green Santolina ", type text}, {"Sage", type text}, {"Pink Texas Skullcap", type text}, {"Rock Rose", type text}, {"Four Nerve Daisy", type text}, {"Fall Aster", type text}, {"Autumn Sage", type text}, {"Gregg's Mistflower", type text}, {"Prairie Verbena", type text}, {"Yarrow", type text}, {"Shasta Daisy", type text}, {"Turk's Cap", type text}, {"Lantana", type text}, {"Additional Comments2", type text}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"ID"}, #"Agg Weather", {"ID"}, "Agg Weather", JoinKind.LeftOuter),
    #"Expanded Agg Weather" = Table.ExpandTableColumn(#"Merged Queries", "Agg Weather", {"SunIndex"}, {"Agg Weather.SunIndex"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Agg Weather",{"Weather"}),
    #"Unpivoted Columns" = Table.UnpivotOtherColumns(#"Removed Columns", {"ID", "Start time", "Completion time", "Email", "Name", "Name2", "Date", "Time", "Additional Comments2", "Agg Weather.SunIndex"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Columns",{{"Attribute", "Plant Type"}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Renamed Columns", "Value", Splitter.SplitTextByDelimiter(";", QuoteStyle.Csv), {"Value.1", "Value.2", "Value.3", "Value.4", "Value.5", "Value.6"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Value.1", type text}, {"Value.2", type text}, {"Value.3", type text}, {"Value.4", type text}, {"Value.5", type text}, {"Value.6", type text}}),
    #"Unpivoted Columns1" = Table.UnpivotOtherColumns(#"Changed Type1", {"ID", "Start time", "Completion time", "Email", "Name", "Name2", "Date", "Time", "Additional Comments2", "Agg Weather.SunIndex", "Plant Type"}, "Attribute", "Value"),
    #"Removed Columns1" = Table.RemoveColumns(#"Unpivoted Columns1",{"Attribute"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns1", each ([Value] <> "" and [Value] <> "   " and [Value] <> "    " and [Value] <> "     ")),
    #"Renamed Columns1" = Table.RenameColumns(#"Filtered Rows",{{"Value", "Observation"}, {"Date", "Survey Date"}})
in
    #"Renamed Columns1"

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