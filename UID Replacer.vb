Sub UIDs()
'
' UIDs Macro
'
' Keyboard Shortcut: Ctrl+Shift+U
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Replace system assigned UIDs with actual UIDs '''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
    Dim rngA, rngB As Range
    
    Set ws = ActiveSheet
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    i = 2

' Loop for NLP feedback
Do While i <= count
    
    If InStr(1, Cells(i, 5).Value, "user") > 0 And _
        InStr(1, Cells(i, 14).Value, "<span hidden id='uid'>") > 0 Then
            systemUID = Cells(i, 5).Value
            realUID = Left(Right(Cells(i, 14).Value, 13), 5)
            
            Columns("E:E").Select
            Selection.Replace What:=systemUID, Replacement:=realUID, lookat:=xlPart _
            , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    End If
         
    i = i + 1
        
Loop

End Sub