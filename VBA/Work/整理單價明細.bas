Attribute VB_Name = "Module10"
Sub ��z�������()
Attribute ��z�������.VB_ProcData.VB_Invoke_Func = " \n14"
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
    
    Range("AT1") = "����Ƹ�"
    Range("AU1") = "���誩��"
    Range("AV1") = "�q��Ƹ�"
    Range("AW1") = "�q�檩��"
    Range("AX1") = "�̥��誩�����D"
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    Range("AT2").Select
    ActiveCell.Formula = "=I2"
    Selection.AutoFill Destination:=Range("AT2:AT" & lrow)
    
    
    Range("AU2").Select
    ActiveCell.Formula = "=J2"
    Selection.AutoFill Destination:=Range("AU2:AU" & lrow)
    
    Range("AV2").Select
    ActiveCell.Formula = "=IF(N2="""","""",LEFT(N2,FIND(""#"",N2,1)-1))"
    Selection.AutoFill Destination:=Range("AV2:AV" & lrow)
    
    Range("AW2").Select
    ActiveCell.Formula = "=IF(N2="""","""",MID(N2,FIND(""#"",N2,1)+1,5))"
    Selection.AutoFill Destination:=Range("AW2:AW" & lrow)
    
    Range("AX2").Select
    ActiveCell.Formula = "=IFERROR(INDEX(A:A,MATCH(AV2,AT:AT,0),1),"""")"
    Selection.AutoFill Destination:=Range("AX2:AX" & lrow)
    
    ActiveSheet.Range("AX2", ActiveSheet.Range("AX" & ActiveSheet.Rows.Count).End(xlUp)).Select
    Selection.Copy
    
    Range("T1").Select
    
    Worksheets("RD�q������X").Activate
    ActiveSheet.Range("AP2").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    

    Range("AP1").Select
    
End Sub
