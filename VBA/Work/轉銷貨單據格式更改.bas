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
    
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "���誩����z"
    
    
    
End Sub



