Sub FixResponse()
'
' FixResponse Macro
'

' Raw data from our NLP Chat Bot comes in on several separate lines for each interaction, one line being the question, another being the response, another being the feedback associated with it. This code locates the question and combines it with the response so that we can see the actual QnA pair for evaluating performance. 

' Clear any active filters
    ActiveSheet.AutoFilterMode = False

' Sort by UserID and then Timestamp
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("C1"), Order:=xlAscending
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SetRange Range("A:AF")
        .Header = xlYes
        .Apply
    End With
    
' Init vars
    Dim count, i As Long
    Dim ws As Worksheet
    Dim cellHolder

    Set ws = ActiveSheet
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    i = 2

' Loop for NLP feedback
Do While i <= count
    
    If Cells(i, 10).Value <> "[]" Then ' If cell shows an intent list...
              
        If Cells(i, 10).Offset(1, 4).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 1 line down")
            cellHolder = Cells(i, 10).Offset(1, 4).Value
            Cells(i, 10).Offset(0, 4).Value = cellHolder
        
        ElseIf Cells(i, 10).Offset(2, 4).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 2 lines down")
            cellHolder = Cells(i, 10).Offset(2, 4).Value
            Cells(i, 10).Offset(0, 4).Value = cellHolder
        
        ElseIf Cells(i, 10).Offset(3, 4).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 3 lines down")
            cellHolder = Cells(i, 10).Offset(3, 4).Value
            Cells(i, 10).Offset(0, 4).Value = cellHolder
            
        ElseIf Cells(i, 10).Offset(4, 4).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 4 lines down")
            cellHolder = Cells(i, 10).Offset(4, 4).Value
            Cells(i, 10).Offset(0, 4).Value = cellHolder
            
        End If

    End If

    i = i + 1
        
Loop

End Sub