Sub 攻牙基線總表_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("攻牙基線總表").Activate
    
    Dim x As Integer
    x = Worksheets("攻牙基線總表").Range("A65536").End(xlUp).Row

    For i = 3 To x
       For j = Range("B1") To Range("HC1")
       
            Cells(i, j).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
                  
            Application.CutCopyMode = False
            
            Range("B3").Select
        
        Next
    Next
      
End Sub

Sub 攻牙基線總表_輸入()
    
    machine_number = Worksheets("攻牙基線表").Range("AA3").Text
    
    '--------------------------------
    
    Worksheets("攻牙基線總表").Activate
    
    Dim x As Integer
    x = Worksheets("攻牙基線總表").Range("A65536").End(xlUp).Row
    
    
    For i = 3 To x
        If Range("A" & i) = machine_number Then
            For j = Range("B1") To Range("HC1") Step 7
                If Cells(i, j) = "" Then
                    Worksheets("攻牙基線表").Range("AB3:AG3").Copy
                    Worksheets("攻牙基線總表").Cells(i, j).PasteSpecial xlPasteValues
                    
                    Worksheets("攻牙基線表").Activate
                    Range("AA3").Select
                    Application.CutCopyMode = False
                    
                    Range("AC3:AD3").ClearContents
                    Range("AF3").ClearContents
                    Range("AE6").ClearContents
                    Range("AG6").ClearContents

                    Exit For
                End If
            Next
        End If
    Next
    
End Sub

