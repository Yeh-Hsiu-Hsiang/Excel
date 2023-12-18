Sub 電鍍發放_DRY001()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-001").Activate
    
    Dim x, y As Integer
    x = Worksheets("電鍍發放").Range("A65536").End(xlUp).Row
    y = Worksheets("DRY-001").Range("A65536").End(xlUp).Row
    
    Workbooks(ActWb).Worksheets("電鍍發放").Activate
    
    For i = 2 To x
        If InStr(Cells(i, "R").Value, "1") > 0 Then
            Range("A" & i & ":U" & i).Copy
        
            Workbooks(ActWb).Worksheets("DRY-001").Activate
            
            j = 4
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

    For i = 5 To y
        For j = Range("A1") To Range("U1")
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
        Next
    Next
    
    Range("A5").Select
    
End Sub

Sub 電鍍發放_DRY002()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-002").Activate
    
    Dim x, y As Integer
    x = Worksheets("電鍍發放").Range("A65536").End(xlUp).Row
    y = Worksheets("DRY-002").Range("A65536").End(xlUp).Row
    
    Workbooks(ActWb).Worksheets("電鍍發放").Activate
    
    For i = 2 To x
        If InStr(Cells(i, "R").Value, "2") > 0 Then
            Range("A" & i & ":U" & i).Copy
        
            Workbooks(ActWb).Worksheets("DRY-002").Activate
            
            j = 4
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

Sub 電鍍發放_DRY002_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-002").Activate
    
    Dim y As Integer
    y = Worksheets("DRY-002").Range("A65536").End(xlUp).Row

    For i = 5 To y
        For j = Range("A1") To Range("U1")
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
        Next
    Next
    
    Range("A5").Select
    
End Sub

Sub 電鍍發放_DRY003()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-003").Activate
    
    Dim x, y As Integer
    x = Worksheets("電鍍發放").Range("A65536").End(xlUp).Row
    y = Worksheets("DRY-003").Range("A65536").End(xlUp).Row
    
    Workbooks(ActWb).Worksheets("電鍍發放").Activate
    
    For i = 2 To x
        If InStr(Cells(i, "R").Value, "3") > 0 Then
            Range("A" & i & ":U" & i).Copy
        
            Workbooks(ActWb).Worksheets("DRY-003").Activate
            
            j = 4
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

Sub 電鍍發放_DRY003_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-003").Activate
    
    Dim y As Integer
    y = Worksheets("DRY-003").Range("A65536").End(xlUp).Row

    For i = 5 To y
        For j = Range("A1") To Range("U1")
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
        Next
    Next
    
    Range("A5").Select
    
End Sub
Sub 電鍍發放_DRY004()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-004").Activate
    
    Dim x, y As Integer
    x = Worksheets("電鍍發放").Range("A65536").End(xlUp).Row
    y = Worksheets("DRY-004").Range("A65536").End(xlUp).Row
    
    Workbooks(ActWb).Worksheets("電鍍發放").Activate
    
    For i = 2 To x
        If InStr(Cells(i, "R").Value, "4") > 0 Then
            Range("A" & i & ":U" & i).Copy
        
            Workbooks(ActWb).Worksheets("DRY-004").Activate
            
            j = 4
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

Sub 電鍍發放_DRY004_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-004").Activate
    
    Dim y As Integer
    y = Worksheets("DRY-004").Range("A65536").End(xlUp).Row

    For i = 5 To y
        For j = Range("A1") To Range("U1")
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
        Next
    Next
    
    Range("A5").Select
    
End Sub
Sub 電鍍發放_DRY005()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-005").Activate
    
    Dim x, y As Integer
    x = Worksheets("電鍍發放").Range("A65536").End(xlUp).Row
    y = Worksheets("DRY-005").Range("A65536").End(xlUp).Row
    
    Workbooks(ActWb).Worksheets("電鍍發放").Activate
    
    For i = 2 To x
        If InStr(Cells(i, "R").Value, "5") > 0 Then
            Range("A" & i & ":U" & i).Copy
        
            Workbooks(ActWb).Worksheets("DRY-005").Activate
            
            j = 4
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

Sub 電鍍發放_DRY005_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("DRY-005").Activate
    
    Dim y As Integer
    y = Worksheets("DRY-005").Range("A65536").End(xlUp).Row

    For i = 5 To y
        For j = Range("A1") To Range("U1")
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
        Next
    Next
    
    Range("A5").Select
    
End Sub
