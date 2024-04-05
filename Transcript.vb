Sub Transcript()
'
' Transcript Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
'

Dim userID, curRow, firstRow, lastRow

If ActiveSheet.Name = "YTD_Integrated" Then
   
    If ActiveCell.Column = 4 Then ' If active cell is USER ID...
        'MsgBox ActiveCell.Column
        Sheets.Add after:=ActiveSheet
        ActiveSheet.Name = "Transcript"
        
        ' Move back to data sheet
        Worksheets(3).Activate
        
        userID = ActiveCell.Value
        curRow = ActiveCell.Row
        firstRow = curRow - 6
        lastRow = curRow + 5
        
        
        Dim source, target As Range
                
        Set source = ActiveSheet.Range("A" & firstRow, "J" & lastRow)
        
        Set target = Sheets("Transcript").Range("A1", "J11")
        
        target.Value = source.Value
              
        Sheets("Transcript").Activate
    
        ' Adjust column widths appropriately
        Columns("B:B").EntireColumn.AutoFit
        Columns("B:B").EntireColumn.AutoFit
        Columns("C:C").EntireColumn.AutoFit
        Columns("D:D").EntireColumn.AutoFit
        Columns("E:E").ColumnWidth = 10
        Columns("F:F").ColumnWidth = 65
        Columns("G:G").ColumnWidth = 65
        Columns("H:H").EntireColumn.AutoFit
        Columns("I:I").EntireColumn.AutoFit
        Columns("J:J").EntireColumn.AutoFit

        ' Filter remaining SuperDA data
        ActiveSheet.Range("$A$1:$J$11").AutoFilter Field:=4, Criteria1:=userID, Operator:=xlFilterValues
        
        ' Remove text wrapping
        Cells.WrapText = False
    
        Rows(1).Hidden = True
        
        MsgBox ("Click OK after reviewing")
        
        Application.DisplayAlerts = False
        Sheets("Transcript").Delete
        Application.DisplayAlerts = True
        
        Sheets("YTD_Integrated").Activate
    
    Else
        MsgBox ("Utilize this function on the User ID field only.")

    End If
Else
    MsgBox ("Can only be process on YTD_Integrated sheet.")

End If

    
End Sub
