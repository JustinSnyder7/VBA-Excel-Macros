Sub xUncheckBoxes()
'
' xUncheckBoxes Macro
'
' Clears the series of check boxes from our testing form.

    Dim chkBox As Excel.CheckBox
    Application.ScreenUpdating = False
    For Each chkBox In ActiveSheet.CheckBoxes
            chkBox.Value = xlOff
    Next chkBox
    Application.ScreenUpdating = True

End Sub