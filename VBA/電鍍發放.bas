Sub 電鍍發放_DRY001()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-001").Activate
    
    Dim x, y As Integer
    x = Worksheets("電鍍發放").Range("A65536").End(xlUp).Row
    y = Worksheets("DRY-001").Range("A65536").End(xlUp).Row
    
    Workbooks(ActWb).Worksheets("電鍍發放").Activate
    
    For i = 2 To x
        If InStr(Cells(i, "R").Value, "1") > 0 Then  '機台包含1
            Range("A" & i & ":U" & i).Copy
        
            Workbooks(ActWb).Worksheets("DRY-001").Activate
            
            j = 3
            Do While True
                If ActiveSheet.Cells(j, 1).Value = "" Then
                    ActiveSheet.Cells(j, 1).Select
                    Exit Do
                End If
                j = j + 1
            Loop
            
            Selection.PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            
            Workbooks(ActWb).Worksheets("電鍍發放").Activate
            
        End If
    Next

End Sub

Sub 電鍍發放_DRY001_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-001").Activate
    
    Dim y As Integer
    y = Worksheets("DRY-001").Range("A65536").End(xlUp).Row

    For i = 3 To y
        For j = Range("A1") To Range("U1")
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
        Next
    Next
    
    Range("A3").Select
    
End Sub

End Sub
