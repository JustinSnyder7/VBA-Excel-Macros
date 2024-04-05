Sub xThumbs()
'
' xThumbs Macro
'
' Keyboard Shortcut: Ctrl+x
'

' Append thumbs up/down to response line

' Clear any active filters
    ActiveSheet.AutoFilterMode = False

' Name column AB:AB to hold thumbs up/down
    Range("$AF$1").Value = "Thumbs"

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
    Dim feedbackUID

    Set ws = ActiveSheet
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    i = 2

' Loop for NLP feedback
Do While i <= count
    
    If Cells(i, 13).Value = "negativeFeedback" And Cells(i, 15).Value = "System.CommonResponse" Or _
        Cells(i, 13).Value = "positiveFeedback" And Cells(i, 15).Value = "System.CommonResponse" Then
        
        feedbackUID = Cells(i, 5).Value
        
        ' If UID is still equal to the same UID check to make sure has a question and answer and also System.Intent
        If Cells(i, 5).Offset(-1, 0).Value = feedbackUID And _
            Cells(i, 13).Offset(-1, 0).Value <> "" And Cells(i, 14).Offset(-1, 0).Value <> "" And _
            Cells(i, 15).Offset(-1, 0).Value = "System.Intent" Then
            
            Cells(i, 32).Offset(-1, 0).Value = Cells(i, 13).Value

        ElseIf Cells(i, 5).Offset(-2, 0).Value = feedbackUID And _
            Cells(i, 13).Offset(-2, 0).Value <> "" And Cells(i, 14).Offset(-2, 0).Value <> "" And _
            Cells(i, 15).Offset(-2, 0).Value = "System.Intent" Then
        
            Cells(i, 32).Offset(-2, 0).Value = Cells(i, 13).Value
        
        ElseIf Cells(i, 5).Offset(-3, 0).Value = feedbackUID And _
            Cells(i, 13).Offset(-3, 0).Value <> "" And Cells(i, 14).Offset(-3, 0).Value <> "" And _
            Cells(i, 15).Offset(-3, 0).Value = "System.Intent" Then
            
            Cells(i, 32).Offset(-3, 0).Value = Cells(i, 13).Value
        
        ElseIf Cells(i, 5).Offset(-4, 0).Value = feedbackUID And _
            Cells(i, 13).Offset(-4, 0).Value <> "" And Cells(i, 14).Offset(-4, 0).Value <> "" And _
            Cells(i, 15).Offset(-4, 0).Value = "System.Intent" Then
            
            Cells(i, 32).Offset(-4, 0).Value = Cells(i, 13).Value
            
        ElseIf Cells(i, 5).Offset(-5, 0).Value = feedbackUID And _
            Cells(i, 13).Offset(-5, 0).Value <> "" And Cells(i, 14).Offset(-5, 0).Value <> "" And _
            Cells(i, 15).Offset(-5, 0).Value = "System.Intent" Then
            
            Cells(i, 32).Offset(-5, 0).Value = Cells(i, 13).Value
            
        End If
    End If
      
        i = i + 1
        
Loop

End Sub