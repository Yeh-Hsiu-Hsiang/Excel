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
    
    
    
    
    '-------------正航版本整理-------------
    Sheets("正航版本整理").Select
    
    Dim lrowT, lrow2 As Long
    lrowT = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    lrow2 = Sheets("轉銷貨單據欄位").Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    
    '-------------正航料號-------------
    Range("C1") = "正航料號"
    Range("C2").Select
    ActiveCell.Formula = "=IF(A2="""","""",LEFT(A2,FIND(""#"",A2,1)-1))"
    
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C" & lrowT)
    '-------------正航料號-------------
    
    
    '-------------正航版本-------------
    Range("D1") = "正航版本"
    Range("D2").Select
    ActiveCell.Formula = "=IF(A2="""","""",MID(A2,FIND(""#"",A2,1)+1,2))"
    
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lrowT)
    '-------------正航版本-------------
    
    
    '-------------訂單料號-------------
    Range("E1") = "訂單料號"
    Range("E2").Select
    ActiveCell.Formula = "=IF(轉銷貨單據欄位!AJ2="""","""",LEFT(轉銷貨單據欄位!AJ2,FIND(""#"",轉銷貨單據欄位!AJ2,1)-1))"
    
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & lrow2)
    '-------------訂單料號-------------
    
    
    '-------------訂單版本-------------
    Range("F1") = "訂單版本"
    Range("F2").Select
    ActiveCell.Formula = "=IF(轉銷貨單據欄位!AJ2="""","""",MID(轉銷貨單據欄位!AJ2,FIND(""#"",轉銷貨單據欄位!AJ2,1)+1,2))"
    
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F" & lrow2)
    '-------------訂單版本-------------
    
    
    '-------------訂單依正航版本為主-------------
    Range("G1") = "訂單依正航版本為主"
    Range("G2").Select
    ActiveCell.Formula = "=SUBSTITUTE(IFERROR(INDEX(A:A,MATCH(E2,C:C,0),1), IF(E2="""","""", E2&""#""&F2)),""#0"",""#O"",1)"
    
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & lrow2)
    
    Range("G2", "G" & Range("G65536").End(xlUp).Row).Select
    Selection.Copy
    
    Sheets("轉銷貨單據欄位").Select
    Range("AJ2").Select
    Selection.PasteSpecial xlPasteValues
    '-------------訂單依正航版本為主-------------
    
    Application.CutCopyMode = False

    '-------------正航版本整理-------------
    

    'Workbooks.Open Filename:="C:\Users\candy\Desktop\TEST銷貨.xls"  '開啟檔案
    Workbooks.Open Filename:="C:\Users\ywqa011\Desktop\雅琪\銷貨\TEST銷貨.xls"  '開啟檔案
    
    Range("A2", "AX" & Range("A65536").End(xlUp).Row + 1).Select
    Selection.Delete
    
    Workbooks(wb).Worksheets("轉銷貨單據欄位").Activate
    
    Range("A2", "AX" & Range("A65536").End(xlUp).Row).Select
    Selection.Copy
    
    Workbooks("TEST銷貨.xls").Worksheets(1).Activate
    
    Range("A2").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    
End Sub



