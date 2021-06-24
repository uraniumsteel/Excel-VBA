Sub CreatePivots()
'Now create the Pivot Table
    Dim pSheet As Worksheet

    Sheets("DataList").Activate
    Set pSheet = Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1", Version:=7).CreatePivotTable TableDestination:=pSheet.Range("A1"), _
        TableName:="PivotTable1", DefaultVersion:=7
    pSheet.Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    Application.WindowState = xlNormal

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Condition")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Days Since Last Updated"), _
        "Sum of Days Since Last Updated", xlSum

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Status")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("RTO Lead Team")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Name")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "Sum of Days Since Last Updated")
        .NumberFormat = "0"
    End With

'Do the Count Pivot Table
pSheet.Select
    Cells("G1").Select
    With ActiveSheet.PivotTables("PivotTable2")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    Application.WindowState = xlNormal

    'ADD ROW IN PIVOT TABLE with "Status" field'
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Status")
        .Orientation = xlRowField
        .Position = 1
    End With
    'ADD ROW BELOW STATUS WITH "RTO Lead Team" Field
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("RTO Lead Team")
        .Orientation = xlRowField
        .Position = 2
    End With
    'ADD ROW BELOW RTO LEAD TEAM WITH "Name" field'
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Name")
        .Orientation = xlRowField
        .Position = 3
    End With
    'Add Values in Pivot Table with Sum  "Count of Lead" field    With ActiveSheet.PivotTables("PivotTable2").PivotFields( _
        "Count of Lead")
        .NumberFormat = "0"
    End With
End Sub
