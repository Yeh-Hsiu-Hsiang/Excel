Sub ��P�f��ڮ榡���()

    Dim wb As String

    wb = ActiveWorkbook.Name
    
    '-------------��P�f������-------------
    
    Sheets("��P�f������").Select
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    
    '-------------#0��#O-------------
    Range("AY1") = "#0��#O"
    Range("AY2").Select
    ActiveCell.Formula = "=SUBSTITUTE(AJ2,""#0"",""#O"",1)"
    
    Range("AY2").Select
    Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
    
    Range("AY2", "AY" & Range("AY65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("AJ2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------#0��#O-------------
    
    
    '-------------�P�f���-------------
    Range("AZ1") = "�P�f���"
    Range("AZ2").Select
    ActiveCell.Formula = "=LEFT(A2,8)"
    
    Range("AZ2").Select
    Selection.AutoFill Destination:=Range("AZ2:AZ" & lrow)

    Range("AZ2", "AZ" & Range("AZ65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("B2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------�P�f���-------------
    
    
    '-------------�Ȥ�s��-------------
    Range("BA1") = "�Ȥ�s��"
    Range("BA2").Select
    Range("BA2:BA" & lrow) = "A00033"
    
    Range("BA2", "BA" & Range("BA65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("C2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------�Ȥ�s��-------------
    
    
    '-------------�~�ȤH��-------------
    Range("BB1") = "�~�ȤH��"
    Range("BB2:BB" & lrow) = "W20020201"
    
    Range("BB2", "BB" & Range("BB65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("G2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------�~�ȤH��-------------
    
    
    '-------------���ݳ���-------------
    Range("BC1") = "���ݳ���"
    Range("BC2:BC" & lrow) = "YW10"
    
    Range("BC2", "BC" & Range("BC65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("H2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------���ݳ���-------------
    
    
    '-------------�ϥι��O-------------
    Range("BD1") = "�ϥι��O"
    Range("BD2").Select
    ActiveCell.Formula = "NTD"
    
    Range("BD2").Select
    Selection.AutoFill Destination:=Range("BD2:BD" & lrow)
    
    Range("BD2", "BD" & Range("BD65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("I2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------�ϥι��O-------------
    
    
    '-------------�ܮw�s��-------------
    Range("BE1") = "�ܮw�s��"
    Range("BE2").Select
    Range("BE2:BE" & lrow) = "YEA03"
   
    Range("BE2", "BE" & Range("BE65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("AL2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------�ܮw�s��-------------
    
    
    Application.CutCopyMode = False
    
    Columns("AY:BE").Select
    Selection.Delete
    '-------------��P�f������-------------
    
    
    
    
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
    

    'Workbooks.Open Filename:="C:\Users\candy\Desktop\TEST�P�f.xls"  '�}���ɮ�
    Workbooks.Open Filename:="C:\Users\ywqa011\Desktop\���X\�P�f\TEST�P�f.xls"  '�}���ɮ�
    
    Range("A2", "AX" & Range("A65536").End(xlUp).Row + 1).Select
    Selection.Delete
    
    Workbooks(wb).Worksheets("��P�f������").Activate
    
    Range("A2", "AX" & Range("A65536").End(xlUp).Row).Select
    Selection.Copy
    
    Workbooks("TEST�P�f.xls").Worksheets(1).Activate
    
    Range("A2").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    
End Sub



