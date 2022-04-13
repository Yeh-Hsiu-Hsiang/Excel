Sub 銷貨_正航版本()

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

End Sub


