Sub 現場沖壓用電基線總表_重設()

    ActWb = ActiveWorkbook.Name
    Workbooks(ActWb).Worksheets("現場沖壓用電基線總表").Activate
    
    Range("C3", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("H3", ActiveSheet.Range("I" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("L3", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("R3", ActiveSheet.Range("R" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("U3", ActiveSheet.Range("X" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("AA3", ActiveSheet.Range("AA" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("AD3", ActiveSheet.Range("AG" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("AJ3", ActiveSheet.Range("AJ" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("AM3", ActiveSheet.Range("AP" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("AS3", ActiveSheet.Range("AS" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("AV3", ActiveSheet.Range("AY" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("BB3", ActiveSheet.Range("BB" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("BE3", ActiveSheet.Range("BH" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("BK3", ActiveSheet.Range("BK" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("BN3", ActiveSheet.Range("BQ" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("BT3", ActiveSheet.Range("BT" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("BW3", ActiveSheet.Range("BZ" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("CC3", ActiveSheet.Range("CC" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("CF3", ActiveSheet.Range("CI" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("CL3", ActiveSheet.Range("CL" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("CO3", ActiveSheet.Range("CR" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("CU3", ActiveSheet.Range("CU" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("CX3", ActiveSheet.Range("DA" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("DD3", ActiveSheet.Range("DD" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("DG3", ActiveSheet.Range("DJ" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("DM3", ActiveSheet.Range("DM" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("DP3", ActiveSheet.Range("DS" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("DV3", ActiveSheet.Range("DV" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
    Range("DY3", ActiveSheet.Range("EB" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    Range("EE3", ActiveSheet.Range("EE" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
    
End Sub

Sub 現場沖壓用電基線總表_輸入()
    
    machine_number = Worksheets("現場沖壓用電基線表單").Range("AE3").Text
    
    '--------------------------------
    
    Worksheets("現場沖壓用電基線總表").Activate
    
    Dim x As Integer
    x = Worksheets("現場沖壓用電基線總表").Range("A65536").End(xlUp).Row   '根據D欄最後一筆資料來找資料共幾列
    
    
    For i = 3 To x
        If Range("A" & i) = machine_number Then
            For j = Range("C1") To Range("JL1") Step 9
                If Cells(i, j) = "" Then
                    Worksheets("現場沖壓用電基線表單").Range("AF3:AI3").Copy
                    Worksheets("現場沖壓用電基線總表").Cells(i, j).PasteSpecial xlPasteValues
                    
                    Worksheets("現場沖壓用電基線表單").Range("AJ3").Copy
                    Worksheets("現場沖壓用電基線總表").Cells(i, j + 6).PasteSpecial xlPasteValues

                    Worksheets("現場沖壓用電基線表單").Activate
                    Range("AE3").Select
                    Application.CutCopyMode = False
                    
                    Range("AH3:AJ3").ClearContents

                    Exit For
                End If
            Next
        End If
    Next
    



    
End Sub
