Sub ��z�������()
'
' ��z������� ����
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
    ActiveWorkbook.Worksheets("���").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("���").Sort.SortFields.Add2 Key:=Range("AN2:AN630") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("���").Sort
        .SetRange Range("AN1:AN630")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("���").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("���").Sort.SortFields.Add2 Key:=Range("AN2:AN630") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("���").Sort
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
    ActiveWorkbook.Worksheets("���").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("���").Sort.SortFields.Add2 Key:=Range("AR3:AR33"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("���").Sort
        .SetRange Range("AR2:AR33")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AR2").Select
    
    
    Range("AP1").Select
    Sheets("���").Select
End Sub



