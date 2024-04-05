Sub toggleAutocalculate()
'
' toggleAutocalculate Macro
'
ActiveWorkbook.PrecisionAsDisplayed = False

With Application
    If .Calculation = xlCalculationManual Then
    .Calculation = xlCalculationAutomatic
Else
    .Calculation = xlCalculationManual
    .MaxChange = 0.001
End If
End With

End Sub