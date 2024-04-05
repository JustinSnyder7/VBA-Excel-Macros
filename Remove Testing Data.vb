Sub xRemoveTestingData()
'
' xDevTest Macro
'
' Remove any data associated with our team of testers.

''''''''''''''''''''''''''' Remove unnecessary button clicks '''''''''''''''''''''''''''
Dim ws As Worksheet
Dim count, i As Long
Dim INTENT

INTENT = 18

Set ws = ActiveSheet

' Set i to start on line 2.
count = ws.Cells(Rows.count, "A").End(xlUp).Row
i = 2

' Loop and paint all intents needing to be deleted a particular color.
Do While i <= count

    If Cells(i, INTENT).Value = "Greeting" _
    Or Cells(i, INTENT).Value = "ExitFlow" _
    Or Cells(i, INTENT).Value = "InvocationJA" _
    Or Cells(i, INTENT).Value = "InvocationPolicy" _
    Or Cells(i, INTENT).Value = "Introduction" _
    Or Cells(i, INTENT).Value = "IcanHelpMenu" _
    Or Cells(i, INTENT).Value = "PayrollIntroduction" _
    Or Cells(i, INTENT).Value = "EmployeeOptionsAP" _
    Or Cells(i, INTENT).Value = "Other Topics" _
    Or Cells(i, INTENT).Value = "fdbck_didntAnswer" _
    Or Cells(i, INTENT).Value = "fdbck_other" _
    Or Cells(i, INTENT).Value = "fdbck_answerConfusing" _
    Or Cells(i, INTENT).Value = "fdbck_notUserFriendly" _
    Or Cells(i, INTENT).Value = "fdbck_thumbsDown" _
    Or Cells(i, INTENT).Value = "fdbck_thumbsUp" _
    Or Cells(i, INTENT).Value = "SupervisorAttendance" _
    Or Cells(i, INTENT).Value = "SupervisorCompensation" _
    Or Cells(i, INTENT).Value = "SupervisorJA" _
    Or Cells(i, INTENT).Value = "SupervisorPayroll" _
    Or Cells(i, INTENT).Value = "SupervisorPolicy" _
    Or Cells(i, INTENT).Value = "SupervisorRole" _
    Or Cells(i, INTENT).Value = "InvocationFMLA" Then
    
    Cells(i, INTENT).Interior.Color = RGB(38, 201, 218)
    
    End If
    
    i = i + 1
    
Loop

ws.Range("A1").AutoFilter Field:=INTENT, Criteria1:=RGB(38, 201, 218), Operator:=xlFilterCellColor

'ws.Cells.Interior.ColorIndex = 0

' Delete the visible rows with the specified color
On Error Resume Next ' If no cells, continues to process
ws.Range("A2:A" & count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
On Error GoTo 0

' Turn off the filter and show all rows
ws.AutoFilterMode = False



''''''''''''''''''''''''' Update UIDs '''''''''''''''''''''''''''

' Temporarily disable automatic calculation so the macro runs quicker.
Application.Calculation = xlManual

' Clear any active filters
    ActiveSheet.AutoFilterMode = False

' Sort by UserID and then Timestamp
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("F1"), Order:=xlAscending
        .SortFields.Add Key:=Range("D1"), Order:=xlAscending
        .SetRange Range("A:AF")
        .Header = xlYes
        .Apply
    End With

' Init vars
    Dim realUID, systemUID
    Dim splitterArray() As String
    Dim rngA, rngB As Range
    
    Set ws = ActiveSheet
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    i = 2

' Loop for NLP feedback
Do While i <= count
    
    If InStr(1, Cells(i, 5).Value, "user") > 0 And _
        InStr(1, Cells(i, 14).Value, "<span hidden id='uid'>") > 0 Then
            systemUID = Cells(i, 5).Value
'            realUID = Left(Right(Cells(i, 14).Value, 13), 6)
            splitterArray = Split(Cells(i, 14), ">", 2)
            realUID = Left(splitterArray(1), 6)
                        
            Columns("E:E").Select
            Selection.Replace What:=systemUID, Replacement:=realUID, lookat:=xlPart _
            , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    End If
         
    i = i + 1
        
Loop

'''''''''''''''''''''''''''''' Clear testing data '''''''''''''''''''''''''''''''

' Set vars
    Set ws = Sheets(1)
    count = ws.Cells(Rows.count, "A").End(xlUp).Row

'
    Rows("1:1").Select
    Selection.AutoFilter


' Filter for team UIDs
    ActiveSheet.Range("$A$1:$AE$1").AutoFilter Field:=5, Criteria1:=Array("js745g", "ja152a", "ms4239", "da243g", "rs1229", "lh7421"), _
        Operator:=xlFilterValues
    
' Delete rows displaying our UIDs
    ActiveSheet.AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)

' Clear filter after delete
    On Error Resume Next
    ActiveSheet.ShowAllData
    
' Filter for any data from the test environment
    ActiveSheet.Range("$A$1:$AE$1").AutoFilter Field:=12, Criteria1:="test", Operator:=xlFilterValues
    
' Delete rows displaying our UIDs
    ActiveSheet.AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)

' Clear filter after delete
    On Error Resume Next
    ActiveSheet.ShowAllData
    
' Rename sheet to SuperDA_Data
    ActiveSheet.Name = "SuperDA_Data"
    
' Re-enable calculation after the macro has run.
Application.Calculation = xlAutomatic
    
End Sub