Sub utterUtterance()
'
' Utter Utterance Macro
'
' This macro is designed to first locate the first row in the Intents file that corresponds with the Intent Name selected in the backLog.
' Once that is determined the corresponding value in the Utterance row of the backLog will automatically paste in the Intents file.
'

' pause screen updating and calculations
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' Initialize variables
Dim utteranceText As String
Dim lines As Variant
Dim utteranceSourceCell As Range
Dim intentSourceCell As Range
Dim pastingCell As Range
Dim TargetCell As Range
Dim targetWorkbook As Workbook
Dim targetWorksheet As Worksheet
Dim matchingRow As Long
Dim cellValue As String
Dim prefix As String


' Check if the active sheet is "Backlog", if not exit sub to avoid any issues.
    If ActiveSheet.Name <> "Backlog" Then
        ' Exit the subroutine if the active sheet is not "Backlog"
        MsgBox "utterUtterance triggered off the Backlog, exiting the subroutine."
        Exit Sub
    End If


' Set references to the target workbook, active cell and worksheet
Set targetWorkbook = Workbooks("SuperDA Backlog.xlsm")

' capture active cell as a range
Set intentSourceCell = ActiveCell
cellValue = intentSourceCell.Value

    ' Check if the cell value contains "_"
    If InStr(cellValue, "_") > 0 Then
        ' Get the prefix
        prefix = Left(cellValue, InStr(cellValue, "_") - 1)
        
        ' Convert the prefix to lowercase
        prefix = LCase(prefix)
        
        ' Check the prefix and do something
        Select Case prefix
            Case "fmla"
                ' Do something for "fmla"
                Set targetWorksheet = ActiveWorkbook.Worksheets("FMLA_Intents")
                ' clear filters from target table
                On Error Resume Next
                    ActiveSheet.ListObjects("fmlaIntents").AutoFilter.ShowAllData
                On Error GoTo 0
                
                
            Case "ja"
                ' Do something for "ja"
                Set targetWorksheet = ActiveWorkbook.Worksheets("JA_Intents")
                
            Case "policy"
                ' Do something for "policy"
                Set targetWorksheet = ActiveWorkbook.Worksheets("Policy_Intents")
                
            Case "payroll"
                ' Do something for "payroll"
                Set targetWorksheet = ActiveWorkbook.Worksheets("Payroll_Intents")
                
            Case "misc"
                ' Do something for "Misc"
                Set targetWorksheet = ActiveWorkbook.Worksheets("Misc_Intents")
                
            Case "ap"
                ' Do something for "Misc"
                Set targetWorksheet = ActiveWorkbook.Worksheets("AP_Intents")
                
            Case Else
                ' Do something else if the prefix is not one of the above
                MsgBox "Skill was not found."
            ' Exit the subroutine
        Exit Sub
        End Select
    Else
        ' Do something if the cell value does not contain "_"
        MsgBox "Skill not found. Cell value does not contain _"
        ' Exit the subroutine
        Exit Sub
    End If
    


' Set targetWorksheet = targetWorkbook.ActiveSheet

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Locate the first row of a given intent name in the current active sheet on the intents file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' grab row number of that corresponds with the intent name in the active cell
On Error Resume Next ' In case no match is found
    matchingRow = Application.WorksheetFunction.match(intentSourceCell.Value, targetWorksheet.Range("B:B"), 0)
    On Error GoTo 0 ' Reset error handling


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Set utteranceText as the cell to the left of the intentSourceCell
Set utteranceSourceCell = intentSourceCell.Offset(0, -1)
    utteranceText = utteranceSourceCell.Value

' Calculate the number of lines in the active cell's value and display it in a message box
    lines = Split(utteranceText, vbLf)
    linesCount = UBound(lines) - LBound(lines) + 1
'    MsgBox (linesCount) ' for diagnostic purposes
    

' Prompt the user to enter the target cell location
Dim cellLocation As Range
Dim userInput As String
On Error Resume Next
    'userInput = InputBox("Please enter the row number:", "Row Number")
    userInput = matchingRow + 1
    If userInput = "" Then Exit Sub
    Set cellLocation = Range("A" & userInput)
    On Error GoTo 0
    
' Validate the entered cell location and display a message if it's invalid
If cellLocation Is Nothing Then
    MsgBox "Invalid cell location. Please try again.", vbExclamation
    Exit Sub
End If

' Set the target cell and insert new rows based on the number of lines in the active cell's value
Set TargetCell = targetWorksheet.Range(cellLocation.Address)
TargetCell.Resize(linesCount).EntireRow.Insert

Set pastingCell = targetWorksheet.Range(cellLocation.Address)

' Inform the user of the selected cell location
' MsgBox "Pasting text in the following cell: " & pastingCell.Address, vbInformation ' for diagnostic purposes

' Paste each line into a separate cell
For i = 0 To linesCount - 1
    pastingCell.Offset(i, 0).Value = lines(i)
Next i

' Move the cursor to the pasting cell location
    Application.ScreenUpdating = False
    targetWorkbook.Activate
    targetWorksheet.Activate

' Use Goto to select and navigate to the cell
    Application.Goto Reference:=targetWorksheet.Range(cellLocation.Offset(-3, 0).Address), Scroll:=True

    Application.ScreenUpdating = True  ' Turn on screen updating

' restore screen updating and calculations
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
    
End Sub