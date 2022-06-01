Attribute VB_Name = "最新資料"
Sub 最新資料()

    Sheets("客戶主檔").Select
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Range("G2").Formula = "=LEFT(H2, FIND(""#"", H2)-1)"
    Range("G2").Select
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    Selection.AutoFill Destination:=Range("G2:G" & lrow)

    For i = 2 To Range("A65536").End(xlUp).Row
        
        If Range("G" & i) = Range("G" & i).Offset(-1, 0) And Range("G" & i) <> "" Then
        
            Rows(i - 1).Delete Shift:=xlUp
            i = i - 1
        End If
    Next
    
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft

End Sub
