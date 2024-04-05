Sub addNewEntry()
'
' addNewEntry Macro
'
' Add new Backlog entry and automate parts of the process where possible

'
' initiate variables
    Dim utter1, intent1, convers1, answer1
    Dim autoConvers, autoAnswer, chkValid
    Dim dateToday
    dateToday = Date


' pull conversation name and store as autoConvers
    If Range("E6").Value = "x" Then
        autoConvers = Application.WorksheetFunction.VLookup(Range("D6"), Range("Intents!C:E"), 2, False)
        Range("E6").Value = autoConvers
    End If

    
' pull answer and store as autoAnswer
    If Range("F6").Value = "x" Then
        autoAnswer = Application.WorksheetFunction.VLookup(Range("D6"), Range("Intents!C:E"), 3, False)
        Range("F6").Value = autoAnswer
        Range("F6").WrapText = False
    End If
    
' grab cell contents from entry point
    utter1 = Range("C6").Value
    intent1 = Range("D6").Value
    convers1 = Range("E6").Value
    answer1 = Range("F6").Value

' insert new row below table header
    Rows(11).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

' paste grabbed information from entry point into new row
    Range("C11").Value = utter1
    Range("D11").Value = intent1
    Range("E11").Value = convers1
    Range("F11").Value = answer1

' auto populate date added
    Range("A11").Value = dateToday

' auto populate Status column to show pending
    Range("H11").Value = "| Pending"
    
' auto populate Reason column to show blank
    Range("G11").Value = "| "
    
' reset entry point for next item
    Range("C6").ClearContents
    Range("D6").ClearContents
    Range("E6").ClearContents
    Range("F6").ClearContents

' turn off word wrapping on answer column
    Range("F11").WrapText = False


End Sub