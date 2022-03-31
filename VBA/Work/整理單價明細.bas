
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
    
    Range("AT1") = "����Ƹ�"
    Range("AU1") = "���誩��"
    Range("AV1") = "�q��Ƹ�"
    Range("AW1") = "�q�檩��"
    Range("AX1") = "OSP�Ƹ�"
    Range("AY1") = "OSP����"
    Range("AZ1") = "�q��̥��誩�����D"
    Range("BA1") = "OSP�̥��誩�����D"
    
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
    ActiveCell.Formula = "=IF(OSP�ॿ����!AR2="""","""",LEFT(OSP�ॿ����!AR2,FIND(""#"",OSP�ॿ����!AR2,1)-1))"
    Selection.AutoFill Destination:=Range("AX2:AX" & lrow)
    
    Range("AY2").Select
    ActiveCell.Formula = "=IF(OSP�ॿ����!AR2="""","""",MID(OSP�ॿ����!AR2,FIND(""#"",OSP�ॿ����!AR2,1)+1,5))"
    Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
    
    
    Range("AZ2").Select
    ActiveCell.Formula = "=SUBSTITUTE(IFERROR(INDEX(A:A,MATCH(AV2,AT:AT,0),1), IF(AV2="""","""", AV2&""#""&AW2)),""#0"",""#O"",1)"
    Selection.AutoFill Destination:=Range("AZ2:AZ" & lrow)
    
    Range("BA2").Select
    ActiveCell.Formula = "=SUBSTITUTE(IFERROR(INDEX(A:A,MATCH(AX2,AT:AT,0),1), IF(AX2="""","""", AX2&""#""&AY2)),""#0"",""#O"",1)"
    Selection.AutoFill Destination:=Range("BA2:BA" & lrow)
    
    ActiveSheet.Range("AZ2", ActiveSheet.Range("AZ" & ActiveSheet.Rows.Count).End(xlUp)).Select
    Selection.Copy
    
    Worksheets("RD�q������X").Activate
    ActiveSheet.Range("AP2").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Worksheets("���").Activate
    ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
    Selection.Copy

    Worksheets("RD�q������X").Activate
    ActiveSheet.Range("AP501").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    Range("AP1").Select
    Sheets("���").Select
End Sub


