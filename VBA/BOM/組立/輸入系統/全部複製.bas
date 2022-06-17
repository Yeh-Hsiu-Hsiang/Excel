Attribute VB_Name = "全部複製"
Sub 全部複製()

    '把原本的資料複製到對應的修改欄位上
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
    
    
    '把原本的資料複製到對應的修改欄位上
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
    
    '------------複製BOM、成品圖、FA------------
    Range("D116:F116").Select
    Selection.Copy
    Range("D116:F116").Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    '------------複製BOM、成品圖、FA------------
    
    
    '------------複製零件圖------------
    For k = 4 To 16
        '------------複製零件圖1~10------------
        Cells(120, k).Select
        Selection.Copy
        Cells(120, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製零件圖1~10------------
        
        
        '------------複製日期版本1~10------------
        Cells(123, k).Select
        Selection.Copy
        Cells(123, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製日期版本1~10------------


        '------------複製零件圖11~20------------
        Cells(127, k).Select
        Selection.Copy
        Cells(127, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製零件圖11~20------------


        '------------複製日期版本11~20------------
        Cells(130, k).Select
        Selection.Copy
        Cells(130, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製日期版本11~20------------


        '------------複製零件圖21~30------------
        Cells(134, k).Select
        Selection.Copy
        Cells(134, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製零件圖21~30------------


        '------------複製日期版本21~30------------
        Cells(137, k).Select
        Selection.Copy
        Cells(137, k).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製日期版本21~30------------
    Next
    '------------複製零件圖------------
    
    
    
    '------------複製成品------------
    Range("D143").Select
    Selection.Copy
    Range("D143").Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    '------------複製成品------------
    
    
    '------------複製零件------------
    For l = 4 To 16
        '------------複製零件1~10------------
        Cells(147, l).Select
        Selection.Copy
        Cells(147, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製零件1~10------------
        
        
        '------------複製日期版本1~10------------
        Cells(150, l).Select
        Selection.Copy
        Cells(150, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製日期版本1~10------------


        '------------複製零件11~20------------
        Cells(154, l).Select
        Selection.Copy
        Cells(154, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製零件11~20------------


        '------------複製日期版本11~20------------
        Cells(157, l).Select
        Selection.Copy
        Cells(157, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製日期版本11~20------------


        '------------複製零件21~30------------
        Cells(161, l).Select
        Selection.Copy
        Cells(161, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製零件21~30------------


        '------------複製日期版本21~30------------
        Cells(164, l).Select
        Selection.Copy
        Cells(164, l).Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        '------------複製日期版本21~30------------
    Next
    '------------複製零件------------
    
End Sub
