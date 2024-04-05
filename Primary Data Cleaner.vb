Sub nuCleaner()
'
' nuCleaner Macro
'

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''' Outline, explainer '''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' 1. Make a copy of raw data.
' 1a. Rename skills to proper naming convention
' 2. Remove duplicate menu responses.
' 3. Remove extra data not needed for any functions.
' 4. Break out datetime to date and time, sort accordingly.
' 5. Add response to intent line.
' 6. Thumbs up/down and feedback survey.
' 6a. New feedback component.
' 7. Break data into Menu vs NLP skill sheets.
' 8. Filter for unresolved Intent, copy to new Integrated tab.
' 9. Go back to NLP tab, filter again and then copy to that same new tab.
' 10. For cleaning menu driven data from the DA.
' 11. Add in unresolvedIntents
' 12. Sort, select and copy data

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Temporarily disable automatic calculation so the macro runs quicker.
Application.Calculation = xlManual

' Init vars
    Dim count, i As Long
    Dim ws As Worksheet
    Dim cellHolder, feedbackUID, rng
    Dim BOT_NAME, USER_UTTERANCE, ENTITY_MATCHES, BOT_RESPONSE, COMPONENT_NAME, DOMAIN_USERID, INTENT, THUMBS, INTENT_LIST, offsetVal
    
' Assign initial values
    BOT_NAME = 1
    DOMAIN_USERID = 4
    USER_UTTERANCE = 6
    BOT_RESPONSE = 7
    COMPONENT_NAME = 8
    ENTITY_MATCHES = 9
    INTENT = 11
    THUMBS = 12
    INTENT_LIST = 5
    offsetVal = 2 ' offset to get to BOT_RESPONSE col


' 1. Make a copy of raw data. ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Verify that you're on the correct file, prevents accidentally running this macro on another file.
    If ActiveSheet.Name <> "SuperDA_Data" Then
        MsgBox ("Does not appear to be raw data, if it is, remove testing data first...")
        Exit Sub
    End If

' Clear any active filters
    ActiveSheet.AutoFilterMode = False
   
' Make a backup copy of sheet for use with Deflection macro
    Sheets("SuperDA_Data").Select
    Sheets("SuperDA_Data").Copy after:=Sheets(1)

' Rename sheet to backupCopy
    ActiveSheet.Name = "backupCopy"
    
' <---- Move back to First Sheet
    Worksheets(1).Activate
    

' 1a. Rename skills to proper naming convention '''''''''''''''''''''''''''''''''''''''''''''''''''

    Columns("A:A").Select
    
    Selection.Replace What:="FMLA*", Replacement:="DAFMLA", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Selection.Replace What:="JA*", Replacement:="DAJA", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Selection.Replace What:="Policy*", Replacement:="DAPOLICY", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Selection.Replace What:="COVID*", Replacement:="DACOVID", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="Birth*", Replacement:="DABIRTH", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="AccountsPayable*", Replacement:="DAAP", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="Miscellaneous*", Replacement:="DAMISC", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="Payroll*", Replacement:="DAPAYROLL", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="Super*", Replacement:="SuperDA", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2


' 2. Remove duplicate menu responses.  Filter COVID and Birth Skills ''''''''''''''''''''''''''''
' filter COMPONENT_NAME for System.CommonResponse and System.Text
' filter ENTITY_MATCHES for {} and delete them.

' Filter Menu Skills
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AE$3824").AutoFilter Field:=BOT_NAME, Criteria1:= _
        "=DABIRTH", Operator:=xlOr, Criteria2:="=DACOVID"
    
' Filter COMPONENT_NAME, column 15 in this state before delete.
'    Rows("1:1").Select
'    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AE$3824").AutoFilter Field:=15, Criteria1:= _
        "=System.CommonResponse", Operator:=xlOr, Criteria2:="=System.Text"
    
' Filter ENTITY_MATCHES
'    Rows("1:1").Select
'    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AE$3824").AutoFilter Field:=ENTITY_MATCHES, Criteria1:= _
        "={}"

' Select and delete filtered data
    Dim sh As Worksheet, LstRw As Long

    Set sh = Sheets("SuperDA_Data")
    With sh
        LstRw = .Cells(.Rows.count, "A").End(xlUp).Row
        Set rng = .Range("A2:A" & LstRw).SpecialCells(xlCellTypeVisible)
        rng.EntireRow.Delete
        .AutoFilterMode = False
    End With

' Clear any active filters
    ActiveSheet.AutoFilterMode = False
    
    
' 3. Remove extra data not needed for any functions. ''''''''''''''''''''''''''''''''''''''''''''
    Columns("S:AF").Select
    Selection.Delete Shift:=xlToLeft

    Columns("K:L").Select
    Selection.Delete Shift:=xlToLeft

    Columns("F:I").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("B:C").Select
    Selection.Delete Shift:=xlToLeft
    

' 4. Break out datetime to date and time, sort accordingly. ''''''''''''''''''''''''''''''''''''''
' Add two new columns
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
     
' Delimit TIMESTAMP by " " into 3 columns
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Tab:=False, Semicolon _
        :=False, Comma:=False, Space:=True, Other:=False, FieldInfo:=Array( _
        Array(1, 4), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
    
    Range("$B$1").Value = "DATE"
    Range("$C$1").Value = "TIME"

' Delete column D
    Columns("D").Select
    Selection.Delete Shift:=xlToLeft

' Resize date column to display properly
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("E:E").ColumnWidth = 1.71
    Columns("F:F").ColumnWidth = 40
    Columns("G:G").ColumnWidth = 40

' Sort by Date, UserID and then Time
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SortFields.Add Key:=Range("D1"), Order:=xlAscending
        .SortFields.Add Key:=Range("C1"), Order:=xlAscending
        .SetRange Range("A:K")
        .Header = xlYes
        .Apply
    End With
  
    Range("S1").Select


' 5. Add response to intent line. ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Loop to add response to correct line
    Set ws = ActiveSheet
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    i = 2

Do While i <= count
      
    If Cells(i, INTENT_LIST).Value <> "[]" Then ' If cell shows an intent list...
              
        If Cells(i, INTENT_LIST).Offset(1, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 1 line down")
            cellHolder = Cells(i, INTENT_LIST).Offset(1, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
        
        ElseIf Cells(i, INTENT_LIST).Offset(2, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 2 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(2, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
        
        ElseIf Cells(i, INTENT_LIST).Offset(3, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 3 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(3, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
            
        ElseIf Cells(i, INTENT_LIST).Offset(4, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 4 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(4, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
            
        ElseIf Cells(i, INTENT_LIST).Offset(5, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 4 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(4, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
            
        ElseIf Cells(i, INTENT_LIST).Offset(6, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 4 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(4, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
            
        ElseIf Cells(i, INTENT_LIST).Offset(7, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 4 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(4, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
            
        ElseIf Cells(i, INTENT_LIST).Offset(8, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 4 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(4, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
            
        ElseIf Cells(i, INTENT_LIST).Offset(9, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 4 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(4, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
            
        ElseIf Cells(i, INTENT_LIST).Offset(10, offsetVal).Value <> "" Then
            'MsgBox ("At line " & i & " there is a response 4 lines down")
            cellHolder = Cells(i, INTENT_LIST).Offset(4, offsetVal).Value
            Cells(i, INTENT_LIST).Offset(0, offsetVal).Value = cellHolder
        End If

    End If

    i = i + 1
        
Loop
    
' 6. Thumbs up/down and feedback survey. '''''''''''''''''''''''''''''''''''''''''''''''''''''

' Clear any active filters
    'ActiveSheet.AutoFilterMode = False

' Rename column S to hold thumbs up/down
    'Range("$L$1").Value = "THUMBS"
    
' Rename column T to hold thumbs up/down
    'Range("$M$1").Value = "SURVEY"

' Prep vars for loop.
    'Set ws = ActiveSheet
    'count = ws.Cells(Rows.count, "A").End(xlUp).Row
    'i = 2


    'i = 2

'Do While i <= count
    
    'If Cells(i, INTENT).Value = "fdbck_thumbsDown" Or _
        'Cells(i, INTENT).Value = "fdbck_thumbsUp" Then
        
        'feedbackUID = Cells(i, DOMAIN_USERID).Value
        
        ' If UID is still equal to the same UID check to make sure has a question and answer and also System.Intent
        'If Cells(i, DOMAIN_USERID).Offset(-1, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-1, 0).Value <> "" Then
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-1, 0).Value Then
                'Cells(i, THUMBS).Offset(-1, 0).Value = Cells(i, INTENT).Value
            'End If

        'ElseIf Cells(i, DOMAIN_USERID).Offset(-2, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-2, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-2, 0).Value Then
                'Cells(i, THUMBS).Offset(-2, 0).Value = Cells(i, INTENT).Value
            'End If
        
        'ElseIf Cells(i, DOMAIN_USERID).Offset(-3, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-3, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-3, 0).Value Then
                'Cells(i, THUMBS).Offset(-3, 0).Value = Cells(i, INTENT).Value
            'End If
        
        'ElseIf Cells(i, DOMAIN_USERID).Offset(-4, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-4, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-4, 0).Value Then
                'Cells(i, THUMBS).Offset(-4, 0).Value = Cells(i, INTENT).Value
            'End If
            
        'ElseIf Cells(i, DOMAIN_USERID).Offset(-5, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-5, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-5, 0).Value Then
                'Cells(i, THUMBS).Offset(-5, 0).Value = Cells(i, INTENT).Value
            'End If
            
        'ElseIf Cells(i, DOMAIN_USERID).Offset(-6, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-6, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-6, 0).Value Then
                'Cells(i, THUMBS).Offset(-6, 0).Value = Cells(i, INTENT).Value
            'End If
            
        'ElseIf Cells(i, DOMAIN_USERID).Offset(-7, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-7, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-7, 0).Value Then
                'Cells(i, THUMBS).Offset(-7, 0).Value = Cells(i, INTENT).Value
            'End If
            
        'ElseIf Cells(i, DOMAIN_USERID).Offset(-8, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-8, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-8, 0).Value Then
                'Cells(i, THUMBS).Offset(-8, 0).Value = Cells(i, INTENT).Value
            'End If
            
        'ElseIf Cells(i, DOMAIN_USERID).Offset(-9, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-9, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-9, 0).Value Then
                'Cells(i, THUMBS).Offset(-9, 0).Value = Cells(i, INTENT).Value
            'End If
            
        'ElseIf Cells(i, DOMAIN_USERID).Offset(-10, 0).Value = feedbackUID And _
            'Cells(i, INTENT).Offset(-10, 0).Value <> "" Then
            
            'If Cells(i, USER_UTTERANCE).Value <> Cells(i, USER_UTTERANCE).Offset(-10, 0).Value Then
                'Cells(i, THUMBS).Offset(-10, 0).Value = Cells(i, INTENT).Value
            'End If
            
'        End If
'    End If
      
'        i = i + 1
        
'Loop

' change text to either positiveFeedback or negativeFeedback
'    Columns("L:L").Select
    
'    Selection.Replace What:="fdbck_thumbsUp", Replacement:="positiveFeedback", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

'    Selection.Replace What:="fdbck_thumbsDown", Replacement:="negativeFeedback", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2


    
' 6b. New feedback component. ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Clear any active filters
    ActiveSheet.AutoFilterMode = False

' Rename column S to hold thumbs up/down
    Range("$L$1").Value = "THUMBS"
    
' Rename column T to hold thumbs up/down
    Range("$M$1").Value = "SURVEY"

' Prep vars for loop.
    Set ws = ActiveSheet
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    i = 2


    i = 2

Do While i <= count
    
    If Cells(i, COMPONENT_NAME).Value = "System.SetCustomMetrics" Then
        
        feedbackUID = Cells(i, DOMAIN_USERID).Value
        
        ' If UID is still equal to the same UID check to make sure has a question and answer and also System.Intent
        If Cells(i, DOMAIN_USERID).Offset(-1, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-1, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-1, 0).Value = Cells(i, USER_UTTERANCE).Value


        ElseIf Cells(i, DOMAIN_USERID).Offset(-2, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-2, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-2, 0).Value = Cells(i, USER_UTTERANCE).Value

        
        ElseIf Cells(i, DOMAIN_USERID).Offset(-3, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-3, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-3, 0).Value = Cells(i, USER_UTTERANCE).Value

        
        ElseIf Cells(i, DOMAIN_USERID).Offset(-4, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-4, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-4, 0).Value = Cells(i, USER_UTTERANCE).Value

            
        ElseIf Cells(i, DOMAIN_USERID).Offset(-5, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-5, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-5, 0).Value = Cells(i, USER_UTTERANCE).Value

            
        ElseIf Cells(i, DOMAIN_USERID).Offset(-6, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-6, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-6, 0).Value = Cells(i, USER_UTTERANCE).Value

            
        ElseIf Cells(i, DOMAIN_USERID).Offset(-7, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-7, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-7, 0).Value = Cells(i, USER_UTTERANCE).Value

            
        ElseIf Cells(i, DOMAIN_USERID).Offset(-8, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-8, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-8, 0).Value = Cells(i, USER_UTTERANCE).Value

            
        ElseIf Cells(i, DOMAIN_USERID).Offset(-9, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-9, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-9, 0).Value = Cells(i, USER_UTTERANCE).Value

            
        ElseIf Cells(i, DOMAIN_USERID).Offset(-10, 0).Value = feedbackUID And _
            Cells(i, INTENT).Offset(-10, 0).Value <> "" Then
                Cells(i, THUMBS).Offset(-10, 0).Value = Cells(i, USER_UTTERANCE).Value

            
        End If
    End If
      
        i = i + 1
        
Loop

' change text to either positiveFeedback or negativeFeedback
    Columns("L:L").Select
    
    Selection.Replace What:="fdbck_thumbsUp", Replacement:="positiveFeedback", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Selection.Replace What:="fdbck_thumbsDown", Replacement:="negativeFeedback", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    
' 7. Break data into Menu vs NLP skill sheets. '''''''''''''''''''''''''''''''''''''''''''''''''
    
' Filter Menu Skills
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AE$3824").AutoFilter Field:=1, Criteria1:= _
        "=DABIRTH", Operator:=xlOr, Criteria2:="=DACOVID"

' Select data from sheet and copy
    Set rng = ActiveSheet.UsedRange
    Intersect(rng, rng.Offset(0)).Copy

    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    
' Rename sheet to Daily MenuDA
    ActiveSheet.Name = "Menu Data"

' Select all cells, just looks cleaner
    Range("A1").Select


' <------ Move back to First Sheet
    Worksheets(1).Activate


' Back at first sheet - unselect previous copied data
    Application.CutCopyMode = False
    
' Now filter for NLP DAs plus SuperDA
    ActiveSheet.Range("$A$1:$AE$3824").AutoFilter Field:=1
    ActiveSheet.Range("$A$1:$AE$3824").AutoFilter Field:=1, Criteria1:=Array( _
        "SuperDA", "DAFMLA", "DAJA", "DAMISC", "DAPOLICY", "DAPAYROLL", "Initial", "DAAP"), Operator:=xlFilterValues

' Select data from sheet and copy
    Set rng = ActiveSheet.UsedRange
    Intersect(rng, rng.Offset(0)).Copy


    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    
' Rename sheet to Daily MenuDA
    ActiveSheet.Name = "NLP Data"

' Select all cells, just looks cleaner
    Range("A1").Select
    
    Set ws = Sheets("NLP Data")
    count = ws.Cells(Rows.count, "A").End(xlUp).Row


' 8. Filter unresolvedIntents from INTENT, copy to new Integrated tab. '''''''''''''''''''''''''''''''''''''''
'    ActiveSheet.Range("$A$1:$H$45").AutoFilter Field:=8, Criteria1:="unresolvedIntent", Operator:=xlFilterValues
   
'    Range("A1").Select
     
' Select and copy cleaned data
'    Set rng = ActiveSheet.UsedRange
'    Intersect(rng, rng.Offset(1)).Copy
    
' add new sheet
    Sheets.Add after:=ActiveSheet
'    ActiveSheet.Paste
    
' Rename sheet to Integrated
   ActiveSheet.Name = "Integrated"
 
' declare variable to count rows
    Dim LR As Long
    LR = Cells(Rows.count, 1).End(xlUp).Row

' Move back to NLP sheet
    Worksheets(2).Activate
    

' 9. Go back to NLP tab, filter again and then copy to that same new tab. ''''''''''''''''''''''
' Utilize color filtering to mark out multiple conditions from INTENT column
i = 2

Do While i <= count

    If Cells(i, INTENT).Value <> "Greeting" _
    And Cells(i, INTENT).Value <> "ExitFlow" _
    And Cells(i, INTENT).Value <> "InvocationJA" _
    And Cells(i, INTENT).Value <> "InvocationPolicy" _
    And Cells(i, INTENT).Value <> "Introduction" _
    And Cells(i, INTENT).Value <> "IcanHelpMenu" _
    And Cells(i, INTENT).Value <> "PayrollIntroduction" _
    And Cells(i, INTENT).Value <> "Other Topics" _
    And Cells(i, INTENT).Value <> "fdbck_didntAnswer" _
    And Cells(i, INTENT).Value <> "fdbck_other" _
    And Cells(i, INTENT).Value <> "fdbck_answerConfusing" _
    And Cells(i, INTENT).Value <> "fdbck_notUserFriendly" _
    And Cells(i, INTENT).Value <> "fdbck_thumbsDown" _
    And Cells(i, INTENT).Value <> "fdbck_thumbsUp" _
    And Cells(i, INTENT).Value <> "SupervisorAttendance" _
    And Cells(i, INTENT).Value <> "SupervisorCompensation" _
    And Cells(i, INTENT).Value <> "SupervisorJA" _
    And Cells(i, INTENT).Value <> "SupervisorPayroll" _
    And Cells(i, INTENT).Value <> "SupervisorPolicy" _
    And Cells(i, INTENT).Value <> "SupervisorRole" _
    And Cells(i, INTENT).Value <> "InvocationFMLA" Then
    
    Cells(i, INTENT).Interior.Color = RGB(38, 201, 218)
    
    End If
    
    i = i + 1
    
Loop

ws.Range("A1").AutoFilter Field:=INTENT, Criteria1:=RGB(38, 201, 218), Operator:=xlFilterCellColor

ws.Cells.Interior.ColorIndex = 0


' Filter out blanks from INTENT_LIST column
    ActiveSheet.Range("$A$1:$I$45").AutoFilter Field:=INTENT_LIST, Criteria1:="<>[]", Operator:=xlFilterValues
    
' Filter remaining SuperDA data
    ActiveSheet.Range("$A$1:$I$45").AutoFilter Field:=BOT_NAME, Criteria1:="<>SuperDA", Operator:=xlFilterValues
        
' Select and copy cleaned data
    Set rng = ActiveSheet.UsedRange
    Intersect(rng, rng.Offset(1)).Copy
     
    
' Move to new Integrated sheet and paste to the end
    Worksheets(3).Activate
    ActiveSheet.Paste Destination:=Worksheets("Integrated").Range("A" & LR + 1)
  
    Range("S1").Select

' Move to Menu sheet
    Worksheets(4).Activate
    
    
' 10. For cleaning menu driven data from the DA. ''''''''''''''''''''''''''''''''''''''''

' Init vars
    Set ws = Sheets("Menu Data")
    count = ws.Cells(Rows.count, "A").End(xlUp).Row

' Remove text wrapping
    Cells.WrapText = False

' Apply filters across header
    Range("A1:M1").Select
    Selection.AutoFilter
   
   
' Utilize color filtering to mark out multiple conditions in USER_UTTERANCE
i = 2

Do While i <= count

    If Cells(i, USER_UTTERANCE).Value <> "Greeting" _
    And Cells(i, USER_UTTERANCE).Value <> "" _
    And Cells(i, USER_UTTERANCE).Value <> "ask BirthRb InvokeMenuBirth" _
    And Cells(i, USER_UTTERANCE).Value <> "negativeFeedback" _
    And Cells(i, USER_UTTERANCE).Value <> "positiveFeedback" _
    And Cells(i, USER_UTTERANCE).Value <> "ask DA-COVID, COVID" Then
    
    Cells(i, USER_UTTERANCE).Interior.Color = RGB(38, 201, 218)
    
    End If
    
    i = i + 1
    
Loop

ws.Range("A1").AutoFilter Field:=USER_UTTERANCE, Criteria1:=RGB(38, 201, 218), Operator:=xlFilterCellColor

ws.Cells.Interior.ColorIndex = 0
   
   
' Filter COMPONENT_NAME
    ActiveSheet.Range("$A$1:$M$85").AutoFilter Field:=COMPONENT_NAME, Criteria1:="System.CommonResponse", Criteria2:="System.Text", Operator:=xlOr
   
' Copy used range on active sheet
    Set rng = ActiveSheet.UsedRange

    Intersect(rng, rng.Offset(1)).Copy

' add new sheet
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste

' Rename sheet to Menu
    ActiveSheet.Name = "Menu"

' Remove column E and then insert a replacement
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

' Copy used range on active sheet
    Set rng = ActiveSheet.UsedRange
    Intersect(rng, rng).Copy

' Move back to Integrated tab
    Worksheets(3).Activate
    
' Init vars
    Set ws = Sheets("Integrated")
    count = ws.Cells(Rows.count, "A").End(xlUp).Row
    
' Init vars to count rows for Integrated sheet
    LR = Cells(Rows.count, 1).End(xlUp).Row

' Paste at the end of Integrated sheet
    ActiveSheet.Paste Destination:=Worksheets("Integrated").Range("A" & LR + 1)

' Remove unneeded columns
    Columns("H:J").Select

    Selection.Delete Shift:=xlToLeft

' 11. Add in unresolvedIntents ''''''''''''''''''''''''''''''''''''''''

' Make a second backup copy of sheet
    Sheets("backupCopy").Select
    Sheets("backupCopy").Copy after:=Sheets(5)

' Rename sheet to backupCopy
    ActiveSheet.Name = "unresolvedIntents"

' 1a. Rename SuperDA to DAMISC '''''''''''''''''''''''''''''''''''''''''''''''''''

    Columns("A:A").Select
    
    Selection.Replace What:="super*", Replacement:="DAMISC", lookat:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2


'    Sheets("unresolvedIntents").Activate
    
    ' Delete columns not needed
    Columns("S:AE").Select
    Selection.Delete Shift:=xlToLeft

    Columns("O:Q").Select
    Selection.Delete Shift:=xlToLeft

    Columns("K:L").Select
    Selection.Delete Shift:=xlToLeft

    Columns("F:I").Select
    Selection.Delete Shift:=xlToLeft

    Columns("B:C").Select
    Selection.Delete Shift:=xlToLeft

' Remove text wrapping
    Cells.WrapText = False
    
    
' Add three new columns
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
''    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
     
' Delimit TIMESTAMP by " " into 3 columns
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Tab:=False, Semicolon _
        :=False, Comma:=False, Space:=True, Other:=False, FieldInfo:=Array( _
        Array(1, 4), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
    
    Range("$B$1").Value = "DATE"
    Range("$C$1").Value = "TIME"
    

' Delete columns C and D
    Columns("D").Select
    Selection.Delete Shift:=xlToLeft

' Resize date column to display properly
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("E:E").ColumnWidth = 20.71
    

' Filter for unresolvedIntent from INTENT
    ActiveSheet.Range("$A$1:$I$45").AutoFilter Field:=8, Criteria1:="unresolvedIntent", Operator:=xlFilterValues
   
    Range("A1").Select
     
' Select and copy cleaned data
'    Dim rng As Range
    Set rng = ActiveSheet.UsedRange
    Intersect(rng, rng.Offset(1)).Copy
    
'   Move back to Integrated sheet
    Worksheets(3).Activate
    
' Set LR to count rows again and paste at the end
    LR = Cells(Rows.count, 1).End(xlUp).Row
    ActiveSheet.Paste Destination:=Worksheets("Integrated").Range("A" & LR + 1)
    
    
    
' 12. Sort, select and copy data ''''''''''''''''''''''''''''''''''''''

' Sort by Date, UserID and then Time
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SortFields.Add Key:=Range("D1"), Order:=xlAscending
        .SortFields.Add Key:=Range("C1"), Order:=xlAscending
        .SetRange Range("A:L")
        .Header = xlNo
        .Apply
    End With

' Remove text wrapping
    Cells.WrapText = False

' Re-enable calculation after the macro has run.
    Application.Calculation = xlAutomatic

' Select and copy cleaned data
    Set rng = ActiveSheet.UsedRange
    Intersect(rng, rng).Copy
    
    
End Sub