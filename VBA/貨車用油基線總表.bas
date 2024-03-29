Sub 貨車用油基線總表_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("貨車用油基線總表").Activate
    
    Dim x As Integer
    x = Worksheets("貨車用油基線總表").Range("A65536").End(xlUp).Row

    For i = 3 To x
       For j = Range("B1") To Range("IG1") Step 8
       
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            

            Cells(i, j + 1).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents

            Cells(i, j + 3).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents

            Cells(i, j + 5).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents

            Cells(i, j + 6).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
                  
            Application.CutCopyMode = False
        
        Next
    Next

End Sub

Sub 貨車用油基線總表_輸入()
    
    number_plate = Worksheets("貨車用油基線表").Range("P4").Text
    
    '--------------------------------
    
    Worksheets("貨車用油基線總表").Activate
    
    Dim x As Integer
    x = Worksheets("貨車用油基線總表").Range("A65536").End(xlUp).Row
    
    
    For i = 3 To x
        If Range("A" & i) = number_plate Then
            For j = Range("B1") To Range("IG1") Step 8
                If Cells(i, j) = "" Then
                    Worksheets("貨車用油基線表").Range("Q4:R4").Copy
                    Worksheets("貨車用油基線總表").Cells(i, j).PasteSpecial xlPasteValues
                    
                    Worksheets("貨車用油基線表").Range("S4").Copy
                    Worksheets("貨車用油基線總表").Cells(i, j + 3).PasteSpecial xlPasteValues
                    
                    Worksheets("貨車用油基線表").Range("T4:U4").Copy
                    Worksheets("貨車用油基線總表").Cells(i, j + 5).PasteSpecial xlPasteValues

                    Worksheets("貨車用油基線表").Activate
                    Range("P4").Select
                    Application.CutCopyMode = False
                    
                    Range("R4:T4").ClearContents
                    Range("U6").ClearContents

                    Exit For
                End If
            Next
        End If
    Next
    
End Sub
