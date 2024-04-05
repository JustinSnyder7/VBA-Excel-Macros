Sub urlExtractor()
'
' urlExtractor Macro
'

Dim IdealMaximum
Dim regEx As Object
Dim curCell As Range

Set regEx = CreateObject("VBScript.RegExp")
IdealMaximum = 380
      
With regEx
        .Pattern = "(http|ftp|https):\/\/([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:\/~+#-]*[\w@?^=%&\/~+#-])"
        .Global = True
        Range("P100").Value = .Replace(ActiveCell.Value, "")
End With
    
' determine if cell length is ideal or not
    If Len(Range("P100").Value) > IdealMaximum Then
        Range("D2").Value = "Queried Cell Length: " & Len(Range("P100").Value)
        Range("D2").Characters(22).Font.Color = vbRed
        Range("P100").Value = ""
    Else
        Range("D2").Value = "Queried Cell Length: " & Len(Range("P100").Value)
        Range("D2").Characters(22).Font.ColorIndex = 10 ' color value for dark green
        Range("P100").Value = ""
    End If

End Sub