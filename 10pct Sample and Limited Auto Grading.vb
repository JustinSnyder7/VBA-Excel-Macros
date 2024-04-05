Sub RandoGordo()
'
' RandoGordo Macro
'
 
' Used to automate daily reports that go to the international team for review.
 
' Table of Contents
' 1 - Auto remove Covid and Birth Skill intents
' 2 - Take 10% randomized sample
' 3 - Change INTENT LIST to be a pure CONFIDENCE score
' 4 - Update certain values automatically based on well understood situations

 
    ' set variables
    Dim rowNum As Long
    Dim topTenPercent As Long
    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range
    Dim confidenceThreshold As Double
    
    ' set the confidence threshold
    confidenceThreshold = 0.98
    
    ' Set reference to the table
    Set tbl = ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1")
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1.''''''''''''''''''''''''''''''''''''''''''''''''''''' Auto remove Covid and Birth Skill intents '''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' remove Covid and Birth skills from the data
    Set rng = tbl.ListColumns("SKILL").DataBodyRange
    For i = rng.count To 1 Step -1 ' Loop backwards since we're deleting rows
        If rng.Cells(i, 1).Value = "DACOVID" Or rng.Cells(i, 1).Value = "DABIRTH" Then
            rng.Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2.'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Take 10% randomized sample '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' create helper column for random number
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=RAND()"
    Range("I3").Select
    Application.Calculation = xlManual 'turn calculation off
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[[#All],[Column1]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' delete used helper column
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    
'    Application.Calculation = xlAutomatic 'turn calculation back on

' take 10% percent of total data and delete the rest
    ' Count the total number of rows
    rowNum = tbl.ListRows.count
    ' Calculate the top 10%
    topTenPercent = Round(rowNum * 0.1)
    ' Delete rows past the top 10%
    If rowNum > topTenPercent Then
        tbl.Range.Rows(topTenPercent + 1 & ":" & rowNum).Delete
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'3.'''''''''''''''''''''''''''''''''''''''''''''''' Change INTENT LIST to be a pure CONFIDENCE score '''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' remove text from INTENT LIST column and leave only primary confidence score
    ' start by removing the value after the "},{" and the first colon
    Set rng = tbl.ListColumns("INTENT LIST").DataBodyRange
    
    For Each cell In rng
        pos = InStr(cell.Value, "},{") ' Find the position of the characters "},{"
        If pos > 0 Then ' If the characters were found
            temp = Left(cell.Value, pos - 1) ' Keep only the text before the characters
            firstColon = InStr(InStr(temp, ":"), temp, ":") ' Find the position of the second colon
            If firstColon > 0 Then ' If the second colon was found
                cell.Value = Mid(temp, firstColon + 1) ' Keep only the text after the second colon
            End If
        End If
    Next cell
           
    ' remove the value after the second colon
    For Each cell In rng
        secondColon = InStr(cell.Value, ":") ' Find the position of the first colon
        If secondColon > 0 Then ' If the first colon was found
            cell.Value = Mid(cell.Value, secondColon + 1) ' Keep only the text after the first colon
        End If
    Next cell
    
    ' Rename the "INTENT LIST" column to "CONFIDENCE"
    tbl.ListColumns("INTENT LIST").Name = "CONFIDENCE"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'4.'''''''''''''''''''''''''''''''''''' Update certain values automatically based on well understood situations ''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' for all instances where CONFIDENCE is greater than an assigned threshold set STATUS as "Bot Action Correct"
    For i = 1 To tbl.ListRows.count
        If tbl.DataBodyRange(i, tbl.ListColumns("CONFIDENCE").Index) >= confidenceThreshold Then
            tbl.DataBodyRange(i, tbl.ListColumns("STATUS").Index) = "Bot Action Correct"
        End If
    Next i

' for all instances where INTENT equals AgentInitation set STATUS as "Agent Transfer"
    For i = 1 To tbl.ListRows.count
        If tbl.DataBodyRange(i, tbl.ListColumns("INTENT").Index) = "AgentInitiation" Then
            tbl.DataBodyRange(i, tbl.ListColumns("STATUS").Index) = "Agent Transfer"
        End If
    Next i

' for all instances where INTENT equals unresolvedIntent set STATUS as "Incorrect"
    For i = 1 To tbl.ListRows.count
        If tbl.DataBodyRange(i, tbl.ListColumns("INTENT").Index) = "unresolvedIntent" Then
            tbl.DataBodyRange(i, tbl.ListColumns("STATUS").Index) = "Incorrect"
        End If
    Next i
    
' for all instances where INTENT equals unresolvedIntent AND confidence score >= .75, set STATUS as "Monitor" and set NOTES as "known issue - double ask"
' the confidence score to resolve an intent is .75, so any that are unresolved but .75 or higher are caused by the client asking for the same intent twice
    For i = 1 To tbl.ListRows.count
        If tbl.DataBodyRange(i, tbl.ListColumns("INTENT").Index) = "unresolvedIntent" And tbl.DataBodyRange(i, tbl.ListColumns("CONFIDENCE").Index) >= 0.75 Then
            tbl.DataBodyRange(i, tbl.ListColumns("STATUS").Index) = "Monitor"
            tbl.DataBodyRange(i, tbl.ListColumns("NOTES").Index) = "known issue - double ask"
        End If
    Next i

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Application.Calculation = xlAutomatic 'turn calculation back on

End Sub