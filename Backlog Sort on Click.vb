Sub refreshStatusSort()
'
' refreshStatusSort Macro - Backlog Status column sort function
'

'
    ActiveWorkbook.Worksheets("Backlog").ListObjects("backlogData").Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("Backlog").ListObjects("backlogData").Sort.SortFields _
        .Add2 Key:=Range("backlogData[Status]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, CustomOrder:= _
        "|- URGENT (notes),| In Progress,| Pending,| w/ SME for Content,| w/ SME for Review,| Completed,| Uploaded,|- Tabled (notes),| Verified Live" _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Backlog").ListObjects("backlogData").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    SendKeys "{F9}"
    
    Application.Wait (Now + TimeValue("0:00:03"))
    
    SendKeys "{NUMLOCK}", True
    
End Sub