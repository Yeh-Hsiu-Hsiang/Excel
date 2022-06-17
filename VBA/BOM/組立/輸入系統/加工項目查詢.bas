Attribute VB_Name = "加工項目查詢"
Sub 加工項目查詢()
    
    '沒有資料會刪除到 F2 標題列
    'Range("F3", Range("F35565").End(xlUp)).Select

    Range("F3:F3100").Clear
    Range("G3:G3100").Clear
    Range("H3:H3100").Clear
    Range("I3:I3100").Clear
    Range("M5:M3100").Clear
    
    For A_Row = 3 To Range("A35565").End(xlUp).Row
    
        '---------------F有資料的最底行---------------
        f = 3
        Do While True
            If ActiveSheet.Cells(f, "F").Value = "" Then
                ActiveSheet.Cells(f, "F").Select
                Exit Do
            End If
            f = f + 1
        Loop
        '---------------F有資料的最底行---------------
        
        '---------------G有資料的最底行---------------
        g = 3
        Do While True
            If ActiveSheet.Cells(g, "G").Value = "" Then
                ActiveSheet.Cells(g, "G").Select
                Exit Do
            End If
            g = g + 1
        Loop
        '---------------G有資料的最底行---------------

        '---------------H有資料的最底行---------------
        h = 3
        Do While True
            If ActiveSheet.Cells(h, "H").Value = "" Then
                ActiveSheet.Cells(h, "H").Select
                Exit Do
            End If
            h = h + 1
        Loop
        '---------------H有資料的最底行---------------

        '---------------I有資料的最底行---------------
        i = 3
        Do While True
            If ActiveSheet.Cells(i, "I").Value = "" Then
                ActiveSheet.Cells(i, "I").Select
                Exit Do
            End If
            i = i + 1
        Loop
        '---------------I有資料的最底行---------------
    
    
    
        '---------------F~I欄資料整理---------------
        If Range("B" & A_Row) <> "" Then
            Range("B" & A_Row).Copy
            Range("F" & f).PasteSpecial xlPasteValues
        End If
        
        
        If Range("C" & A_Row) <> "" Then
            Range("C" & A_Row).Copy
            Range("G" & g).PasteSpecial xlPasteValues
        End If
        
        If Range("D" & A_Row) <> "" Then
            Range("D" & A_Row).Copy
            Range("H" & h).PasteSpecial xlPasteValues
        End If
        
        If Range("E" & A_Row) <> "" Then
            Range("E" & A_Row).Copy
            Range("I" & i).PasteSpecial xlPasteValues
        End If
        '---------------F~I欄資料整理---------------
    Next
    
    Application.CutCopyMode = False
    
    For K_Row = 1 To Range("K35565").End(xlUp).Row
    
        '---------------M有資料的最底行---------------
        m = 5
        Do While True
            If ActiveSheet.Cells(m, "M").Value = "" Then
                ActiveSheet.Cells(m, "M").Select
                Exit Do
            End If
            m = m + 1
        Loop
        '---------------M有資料的最底行---------------
        
        
        If Range("K" & K_Row) <> "" Then
            Range("K" & K_Row).Copy
            Range("M" & m).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If
    Next
    
    MsgBox "查詢完成"

End Sub
