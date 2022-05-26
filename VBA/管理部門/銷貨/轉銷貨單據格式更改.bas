Sub 轉銷貨單據格式更改()

    Dim wb As String

    wb = ActiveWorkbook.Name
    
    '-------------轉銷貨單據欄位-------------
    
    Sheets("轉銷貨單據欄位").Select
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    
    '-------------#0改#O-------------
    Range("AY1") = "#0改#O"
    Range("AY2").Select
    ActiveCell.Formula = "=SUBSTITUTE(AJ2,""#0"",""#O"",1)"
    
    Range("AY2").Select
    Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
    
    Range("AY2", "AY" & Range("AY65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("AJ2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------#0改#O-------------
    
    
    '-------------銷貨日期-------------
    Range("AZ1") = "銷貨日期"
    Range("AZ2").Select
    ActiveCell.Formula = "=LEFT(A2,8)"
    
    Range("AZ2").Select
    Selection.AutoFill Destination:=Range("AZ2:AZ" & lrow)

    Range("AZ2", "AZ" & Range("AZ65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("B2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------銷貨日期-------------
    
    
    '-------------客戶編號-------------
    Range("BA1") = "客戶編號"
    Range("BA2").Select
    Range("BA2:BA" & lrow) = "A00033"
    
    Range("BA2", "BA" & Range("BA65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("C2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------客戶編號-------------
    
    
    '-------------業務人員-------------
    Range("BB1") = "業務人員"
    Range("BB2:BB" & lrow) = "W20020201"
    
    Range("BB2", "BB" & Range("BB65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("G2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------業務人員-------------
    
    
    '-------------所屬部門-------------
    Range("BC1") = "所屬部門"
    Range("BC2:BC" & lrow) = "YW10"
    
    Range("BC2", "BC" & Range("BC65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("H2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------所屬部門-------------
    
    
    '-------------使用幣別-------------
    Range("BD1") = "使用幣別"
    Range("BD2").Select
    ActiveCell.Formula = "NTD"
    
    Range("BD2").Select
    Selection.AutoFill Destination:=Range("BD2:BD" & lrow)
    
    Range("BD2", "BD" & Range("BD65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("I2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------使用幣別-------------
    
    
    '-------------倉庫編號-------------
    Range("BE1") = "倉庫編號"
    Range("BE2").Select
    Range("BE2:BE" & lrow) = "YEA03"
   
    Range("BE2", "BE" & Range("BE65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("AL2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------倉庫編號-------------
    
    
    Application.CutCopyMode = False
    
    Columns("AY:BE").Select
    Selection.Delete
    '-------------轉銷貨單據欄位-------------
    
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "正航版本整理"
    
    
    
End Sub



