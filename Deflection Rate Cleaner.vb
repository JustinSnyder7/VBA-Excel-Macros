Sub deflectionRateClean()
'
' deflectionRateClean Macro
'
' Keyboard Shortcut: Ctrl+d
'

' Delete unneeded columns
    Columns("R:AE").Select
    Selection.Delete Shift:=xlToLeft

    Columns("F:P").Select
    Selection.Delete Shift:=xlToLeft

    Columns("B:C").Select
    Selection.Delete Shift:=xlToLeft
     
' Insert two new columns for process to break up time and date
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
' Change timestamp to date and then also time
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Tab:=False, Semicolon _
        :=False, Comma:=False, Space:=True, Other:=False, FieldInfo:=Array( _
        Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
    
' Delete extraneous column
    Columns("D").Select
    Selection.Delete Shift:=xlToLeft
    
' Rename column to "TIME"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "TIME"
    Range("C2").Select
    
' Filter for "agentConversation"
    ActiveSheet.Range("$A$1:$E$964").AutoFilter Field:=5, Criteria1:="agentConversation", Operator:=xlFilterValues
    
' Copy used range on active sheet including header
    Dim rng

    Set rng = ActiveSheet.UsedRange

    Intersect(rng, rng).Copy
    
' add new sheet
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    
' Rename sheet to Deflection
    ActiveSheet.Name = "Deflection"
    
' set vars
    Dim RowCount As Integer
    
' get count of rows
    RowCount = ActiveSheet.UsedRange.Rows.count

' create pivot table from data on Deflection sheet
    Application.CutCopyMode = False
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Deflection!R1C1:R" & RowCount & "C5", Version:=6).CreatePivotTable TableDestination:= _
        "Deflection!R2C10", TableName:="DeflectionTable1", DefaultVersion:=6
    Sheets("Deflection").Select
    Cells(2, 10).Select
    With ActiveSheet.PivotTables("DeflectionTable1")
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
    With ActiveSheet.PivotTables("DeflectionTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("DeflectionTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("DeflectionTable1").PivotFields("BOT_NAME")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("DeflectionTable1").AddDataField ActiveSheet.PivotTables( _
        "DeflectionTable1").PivotFields("NEXT_STATE"), "Count of NEXT_STATE", xlCount
    
' Copy pertinent information
    Range("K3:K8").Select
    Selection.Copy

    
End Sub