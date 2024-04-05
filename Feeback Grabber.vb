Sub xFeedbackGrabber()
'
' xFeedbackGrabber Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''' For grabbing feedback from the raw data ''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Temporarily disable automatic calculation so the macro runs quicker.
Application.Calculation = xlManual

' Copy backupsheet
    Sheets("backupCopy").Select
    Sheets("backupCopy").Copy before:=Sheets(1)
    
' Rename sheet to backupCopy
    ActiveSheet.Name = "feedbackCopy"

' Init vars
    Dim count, i, j As Long
    Dim ws As Worksheet
    Dim response1, user1, date1, time1, filterVar
    Dim copyRange As String

    Set ws = ActiveSheet
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    definedRows = ws.Cells(Rows.count, "A").End(xlUp).Row

' Clear any active filters before proceeding
    ActiveSheet.AutoFilterMode = False

' Make a copy of the sheet
'    Worksheets("SuperDA_Data").Copy After:=Worksheets("SuperDA_Data")

'''''''''''''''''''''''''' Setup Loop to Remove Agent Data ''''''''''''''''''''''''''''''

' Filter column L for 'serviceCloud' and copy UID's from column B to a new sheet
    ActiveSheet.Range("L:L").AutoFilter Field:=1, Criteria1:="serviceCloud", Operator:=xlFilterValues

' Copy used range on active sheet
    Set rng = ActiveSheet.UsedRange
    Intersect(rng, rng.Offset(1)).Copy

' Paste onto new sheet, rename sheet
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    ActiveSheet.Name = "UIDs"

' Delete columns not needed
    Columns("F:AE").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:D").Select
    Selection.Delete Shift:=xlToLeft

' Reset values
    Set ws = ActiveSheet
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    i = 1


' While still in range
Do While i <= count

    ' assign filterVar to cell
    Set filterVar = ActiveSheet.Cells(i, 1)
    
    ' Move back to first sheet
    Worksheets(1).Activate
    
    ' Clear any active filters before proceeding
    ActiveSheet.AutoFilterMode = False
    
    ' Filter by FilterVar
    ActiveSheet.Range("$A$1:$AE$1").AutoFilter Field:=5, Criteria1:=filterVar.Value, Operator:=xlFilterValues
    
    ' Select and delete filtered data
    ActiveSheet.Range("$A$2:$AE$30000").SpecialCells _
        (xlCellTypeVisible).EntireRow.Delete
    
    ' Clear any active filters before proceeding
    ActiveSheet.AutoFilterMode = False
    
' Move back to second sheet
    Worksheets(2).Activate
    
    i = i + 1
    
Loop

' Move back to first sheet
    Worksheets(1).Activate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Delete columns not needed
    Columns("S:AE").Select
    Selection.Delete Shift:=xlToLeft

    Columns("O:Q").Select
    Selection.Delete Shift:=xlToLeft

    Columns("K:L").Select
    Selection.Delete Shift:=xlToLeft

    Columns("F:I").Select
    Selection.Delete Shift:=xlToLeft

    Columns("A:C").Select
    Selection.Delete Shift:=xlToLeft

' Remove text wrapping
    Cells.WrapText = False
    
    
' Add two new columns
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
     
' Delimit TIMESTAMP by " " into 3 columns
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Tab:=False, Semicolon _
        :=False, Comma:=False, Space:=True, Other:=False, FieldInfo:=Array( _
        Array(1, 4), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
    
    Range("$A$1").Value = "DATE"
    Range("$B$1").Value = "TIME"
    
 
' Delete columns C
    Columns("C").Select
    Selection.Delete Shift:=xlToLeft

' Sort by Date, UserID and then Time
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), Order:=xlAscending
        .SortFields.Add Key:=Range("C1"), Order:=xlAscending
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SetRange Range("A:G")
        .Header = xlNo
        .Apply
    End With
  
    Range("S1").Select

' Resize date column to display properly
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").ColumnWidth = 10
    Columns("D:D").ColumnWidth = 4
    Columns("E:E").ColumnWidth = 40
    Columns("F:F").ColumnWidth = 40
    
' Filter for 'Thanks again, have a great day. Please hold..'
    ActiveWindow.SmallScroll Down:=-36

    Range("A1:G1").Select

    Selection.AutoFilter
        
    ActiveSheet.Range("$A$1:$H$45").AutoFilter Field:=6, Criteria1:="Thanks again, have a great day. Please hold..", Operator:=xlOr, Criteria2:="Before you go, was I helpful today?"



' Filter out blanks from COMPONENT_NAME column
    ActiveSheet.Range("$A$1:$F$45").AutoFilter Field:=5, Criteria1:="<>", Operator:=xlFilterValues
     
    Range("A1").Select
    
    
' Select and copy cleaned data
    Set rng = ActiveSheet.UsedRange
    Intersect(rng, rng.Offset(1)).Copy
    

' add new sheet
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    
' Rename sheet to Throw
    ActiveSheet.Name = "Throw"
    
' Delete column E
    Columns("F").Select
    Selection.Delete Shift:=xlToLeft
    
' Select and copy cleaned data
    Dim rng3 As Range
    Set rng3 = ActiveSheet.UsedRange
    Intersect(rng3, rng3).Copy
    
' add new sheet
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    
' Rename sheet to Feedback
    ActiveSheet.Name = "Feedback"
    
' Delete now unneeded Throw sheet
    Application.DisplayAlerts = False
        Sheets("Throw").Delete
    Application.DisplayAlerts = True

' <----- Move back to SuperDA_Data sheet
    Worksheets(1).Activate
    
' Clear any active filters again
    ActiveSheet.AutoFilterMode = False
    

' Loop to pull data
i = 2
j = 2
Set ws = ActiveSheet
count = ws.Cells(Rows.count, "A").End(xlUp).Row

Do While i <= count

    If Cells(i, 6).Value = "Before you go, was I helpful today?" Then
        
        If Cells(i, 3).Value = Cells(i, 3).Offset(1, 0).Value Then
            date1 = Cells(i, 1).Value
            time1 = Cells(i, 2).Value
            user1 = Cells(i, 3).Value
            response1 = Cells(i, 6).Offset(1, -1).Value
            
    
            Cells(j, 10).Value = date1
            Cells(j, 11).Value = time1
            Cells(j, 12).Value = user1
            Cells(j, 13).Value = response1
            
            j = j + 1
        End If
    
    End If
    
    i = i + 1
    
Loop

Let copyRange = "J2" & ":" & "M" & j

Range(copyRange).Copy

' Move to new Integrated sheet and paste to the end
    Worksheets(2).Activate
    
' Init vars to count rows for Feedback sheet
    Dim LR As Long
    LR = Cells(Rows.count, 1).End(xlUp).Row

' Paste at the end of Feedback sheet
    ActiveSheet.Paste Destination:=Worksheets("Feedback").Range("A" & LR + 1)
    
' Sort by Date, UserID and then Time
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), Order:=xlAscending
        .SortFields.Add Key:=Range("C1"), Order:=xlAscending
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SetRange Range("A:H")
        .Header = xlNo
        .Apply
    End With
    
' Select and copy cleaned data
    Dim rng4 As Range
    Set rng4 = ActiveSheet.UsedRange
    Intersect(rng4, rng4).Copy
    
' Re-enable calculation after the macro has run.
Application.Calculation = xlAutomatic
    
End Sub