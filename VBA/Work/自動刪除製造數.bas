Sub 自動刪除製造數()

    Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

    For i = 6 To ActiveSheet.Range("R65536").End(xlUp).Row
        
        If Range("R" & i) = "" Then
        
            Rows(i).Offset.Select
            Selection.Delete Shift:=xlUp
            
            If Range("R" & i).Offset(-1, 0) = "" Then
                Rows(i).Offset(-1, 0).Select
                Selection.Delete Shift:=xlUp
            End If
        End If
    Next
    
End Sub

