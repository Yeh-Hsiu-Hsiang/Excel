Sub �P�f_���誩��()

    '-------------���誩����z-------------
    Sheets("���誩����z").Select
    
    Dim lrowT, lrow2 As Long
    lrowT = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    lrow2 = Sheets("��P�f������").Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    
    '-------------����Ƹ�-------------
    Range("C1") = "����Ƹ�"
    Range("C2").Select
    ActiveCell.Formula = "=IF(A2="""","""",LEFT(A2,FIND(""#"",A2,1)-1))"
    
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & lrowT)
    '-------------����Ƹ�-------------
    
    
    '-------------���誩��-------------
    Range("D1") = "���誩��"
    Range("D2").Select
    ActiveCell.Formula = "=IF(A2="""","""",MID(A2,FIND(""#"",A2,1)+1,2))"
    
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lrowT)
    '-------------���誩��-------------
    
    
    '-------------�q��Ƹ�-------------
    Range("E1") = "�q��Ƹ�"
    Range("E2").Select
    ActiveCell.Formula = "=IF(��P�f������!AJ2="""","""",LEFT(��P�f������!AJ2,FIND(""#"",��P�f������!AJ2,1)-1))"
    
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & lrow2)
    '-------------�q��Ƹ�-------------
    
    
    '-------------�q�檩��-------------
    Range("F1") = "�q�檩��"
    Range("F2").Select
    ActiveCell.Formula = "=IF(��P�f������!AJ2="""","""",MID(��P�f������!AJ2,FIND(""#"",��P�f������!AJ2,1)+1,2))"
    
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F" & lrow2)
    '-------------�q�檩��-------------
    
    
    '-------------�q��̥��誩�����D-------------
    Range("G1") = "�q��̥��誩�����D"
    Range("G2").Select
    ActiveCell.Formula = "=SUBSTITUTE(IFERROR(INDEX(A:A,MATCH(E2,C:C,0),1), IF(E2="""","""", E2&""#""&F2)),""#0"",""#O"",1)"
    
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & lrow2)
    
    Range("G2", "G" & Range("G65536").End(xlUp).Row).Select
    Selection.Copy
    
    Sheets("��P�f������").Select
    Range("AJ2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------�q��̥��誩�����D-------------
    
    Application.CutCopyMode = False

    '-------------���誩����z-------------

End Sub


