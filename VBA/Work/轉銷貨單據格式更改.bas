Sub 轉銷貨單據格式更改()

    Sheets("轉銷貨單據欄位").Select
    Range("AY1") = "#0改#O"
    Range("AY2").Select
    ActiveCell.Formula = "=SUBSTITUTE(AJ2,""#0"",""#O"",1)"
    
    Range("AZ1") = "銷貨日期"
    Range("AZ2").Select
    ActiveCell.Formula = "=LEFT(A2,8)"
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

    Range("AY2").Select
    Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
    
    Range("AZ2").Select
    Selection.AutoFill Destination:=Range("AZ2:AZ" & lrow)
    
    
    Range("AY2", "AY" & Range("AY65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("AJ2").Select
    Selection.PasteSpecial xlPasteValues
    
    
    Range("AZ2", "AZ" & Range("AZ65536").End(xlUp).Row).Select
    Selection.Copy
    
    Range("B2").Select
    Selection.PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    
    Columns("AY:AZ").Select
    Selection.Delete
    
End Sub
