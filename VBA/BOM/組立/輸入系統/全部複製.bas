Attribute VB_Name = "�����ƻs"
Sub �����ƻs()

    '��쥻����ƽƻs��������ק����W
    For i = 27 To ActiveSheet.Range("C65536").End(xlUp).Row Step 2
        For n = 3 To 26
            Cells(i, n).Select
            Selection.Copy
            Cells(i, n).Offset(1, 0).Select
            Selection.PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            Range("C24").Select
        Next n
    Next i
    
    
    '��쥻����ƽƻs��������ק����W
    For j = 14 To 23 Step 2
        Range("B" & j).Select
        Selection.Copy
        Range("B" & j).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Range("D" & j).Select
        Selection.Copy
        Range("D" & j).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Range("E" & j).Select
        Selection.Copy
        Range("E" & j).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Range("P" & j).Select
        Selection.Copy
        Range("P" & j).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Range("U" & j).Select
        Selection.Copy
        Range("U" & j).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Range("Y" & j).Select
        Selection.Copy
        Range("Y" & j).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        
        Range("B11").Select
    Next j
    
    '------------�ƻsBOM�B���~�ϡBFA------------
    Range("D116:F116").Select
    Selection.Copy
    Range("D116:F116").Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    '------------�ƻsBOM�B���~�ϡBFA------------
    
    
    '------------�ƻs�s���------------
    For k = 4 To 16
        '------------�ƻs�s���1~10------------
        Cells(120, k).Select
        Selection.Copy
        Cells(120, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�s���1~10------------
        
        
        '------------�ƻs�������1~10------------
        Cells(123, k).Select
        Selection.Copy
        Cells(123, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�������1~10------------


        '------------�ƻs�s���11~20------------
        Cells(127, k).Select
        Selection.Copy
        Cells(127, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�s���11~20------------


        '------------�ƻs�������11~20------------
        Cells(130, k).Select
        Selection.Copy
        Cells(130, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�������11~20------------


        '------------�ƻs�s���21~30------------
        Cells(134, k).Select
        Selection.Copy
        Cells(134, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�s���21~30------------


        '------------�ƻs�������21~30------------
        Cells(137, k).Select
        Selection.Copy
        Cells(137, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�������21~30------------
    Next
    '------------�ƻs�s���------------
    
    
    
    '------------�ƻs���~------------
    Range("D143").Select
    Selection.Copy
    Range("D143").Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    '------------�ƻs���~------------
    
    
    '------------�ƻs�s��------------
    For l = 4 To 16
        '------------�ƻs�s��1~10------------
        Cells(147, l).Select
        Selection.Copy
        Cells(147, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�s��1~10------------
        
        
        '------------�ƻs�������1~10------------
        Cells(150, l).Select
        Selection.Copy
        Cells(150, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�������1~10------------


        '------------�ƻs�s��11~20------------
        Cells(154, l).Select
        Selection.Copy
        Cells(154, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�s��11~20------------


        '------------�ƻs�������11~20------------
        Cells(157, l).Select
        Selection.Copy
        Cells(157, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�������11~20------------


        '------------�ƻs�s��21~30------------
        Cells(161, l).Select
        Selection.Copy
        Cells(161, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�s��21~30------------


        '------------�ƻs�������21~30------------
        Cells(164, l).Select
        Selection.Copy
        Cells(164, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------�ƻs�������21~30------------
    Next
    '------------�ƻs�s��------------
    
End Sub
