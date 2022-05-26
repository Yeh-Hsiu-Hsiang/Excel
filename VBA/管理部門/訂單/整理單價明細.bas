Sub 整理單價明細()
'
' 整理單價明細 巨集
'

    Range("T2:T102").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
   
    Range("AN2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$AN$1:$AN$630").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    ActiveWorkbook.Worksheets("單價").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("單價").Sort.SortFields.Add2 Key:=Range("AN2:AN630") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("單價").Sort
        .SetRange Range("AN1:AN630")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("單價").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("單價").Sort.SortFields.Add2 Key:=Range("AN2:AN630") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("單價").Sort
        .SetRange Range("AN1:AN630")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AN2").Select

    Range("AE2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("AR2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("$AR$1:$AR$33").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    ActiveWorkbook.Worksheets("單價").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("單價").Sort.SortFields.Add2 Key:=Range("AR3:AR33"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("單價").Sort
        .SetRange Range("AR2:AR33")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AR2").Select
    
    
    Range("AP1").Select
    Sheets("單價").Select
End Sub



