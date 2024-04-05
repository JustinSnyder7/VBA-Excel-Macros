Sub AgentDriver()
'
' AgentDriver Macro
'
' Component to help determine the sequence of events that led to an agent being needed to assist the client.

'
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim lastRow As Long
    
    ' Set the worksheet and table
    Set ws = Worksheets("Sheet1") 'replace with your sheet name
    Set tbl = ws.ListObjects("ytdIntegrated")
    lastRow = tbl.Range.Rows.count
    ' Loop through each row in the table
    For i = lastRow To 2 Step -1
        ' Check if the intent is "AgentInitiation"
        If tbl.DataBodyRange(i, tbl.ListColumns("INTENT").Index) = "AgentInitiation" Then
            ' Check if the user ID in this row is the same as in the previous row
            If tbl.DataBodyRange(i, tbl.ListColumns("USER ID").Index) = tbl.DataBodyRange(i - 1, tbl.ListColumns("USER ID").Index) Then
                ' Label the previous row
                tbl.DataBodyRange(i - 1, tbl.ListColumns("DRIVER").Index) = "Agent" 'replace with your label column and text
            End If
        End If
    Next i
    
End Sub