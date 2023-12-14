Sub 現場沖壓用電基線總表_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("現場沖壓用電基線總表").Activate
    
    Dim x As Integer
    x = Worksheets("現場沖壓用電基線總表").Range("A65536").End(xlUp).Row

    For i = 3 To x
       For j = Range("C1") To Range("JL1") Step 9
       
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            
            Cells(i, j + 1).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents

            Cells(i, j + 2).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents

            Cells(i, j + 3).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents

            Cells(i, j + 6).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
                  
            Application.CutCopyMode = False
            
            Range("C3").Select
        
        Next
    Next
      
End Sub

Sub 現場沖壓用電基線總表_輸入()
    
    machine_number = Worksheets("現場沖壓用電基線表").Range("AE3").Text
    
    '--------------------------------
    
    Worksheets("現場沖壓用電基線總表").Activate
    
    Dim x As Integer
    x = Worksheets("現場沖壓用電基線總表").Range("A65536").End(xlUp).Row
    
    
    For i = 3 To x
        If Range("A" & i) = machine_number Then
            For j = Range("C1") To Range("JL1") Step 9
                If Cells(i, j) = "" Then
                    Worksheets("現場沖壓用電基線表").Range("AF3:AI3").Copy
                    Worksheets("現場沖壓用電基線總表").Cells(i, j).PasteSpecial xlPasteValues
                    
                    Worksheets("現場沖壓用電基線表").Range("AJ3").Copy
                    Worksheets("現場沖壓用電基線總表").Cells(i, j + 6).PasteSpecial xlPasteValues

                    Worksheets("現場沖壓用電基線表").Activate
                    Range("AE3").Select
                    Application.CutCopyMode = False
                    
                    Range("AH3:AJ3").ClearContents
                    Range("AG6").ClearContents
                    Range("AI6").ClearContents

                    Exit For
                End If
            Next
        End If
    Next
    
End Sub
