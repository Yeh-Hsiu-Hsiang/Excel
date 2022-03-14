Sub Input_GeneralTable()
    
    Dim ActWb As String, i, j, k As Long
    
    ActWb = ActiveWorkbook.Name
    
    For j = 2 To Workbooks(ActWb).Worksheets(1).Range("A65536").End(xlUp).Row '最後一列
    
        '------------選定位置為有資料的最底行------------
        Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
        
        i = 6
        Do While True
            If Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Cells(i, "D").Value = "" Then
                Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Cells(i, "D").Select
                Exit Do
            End If
            i = i + 1
        Loop
        '------------選定位置為有資料的最底行------------
    
        
        
        '----------判斷總表D欄是 首件、IPQC----------
        If InStr(1, ActWb, "首件") <> 0 Then    '判斷是否為首件
            Range("D" & i) = "首件"
        ElseIf InStr(1, ActWb, "QC") <> 0 Then    '判斷是否為IPQC
            Range("D" & i) = "IPQC"
        End If
        '----------判斷總表D欄是 首件、IPQC----------


        Workbooks(ActWb).Worksheets(1).Activate
        For k = 1 To Workbooks(ActWb).Worksheets(1).Cells(1, Columns.Count).End(xlToLeft).Column '最後一欄
        
            '----------檢驗日期----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "QR Code_日期") <> 0 Then   '判斷是否等於日期
                Cells(j, k).Select
                Selection.Copy
            
                Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                Range("E" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------檢驗日期----------
            
            
            '----------檢驗員----------
            Dim inspector As String
            
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "檢驗員") <> 0 Or InStr(1, Cells(1, k), "檢驗者") Then   '判斷是否等於檢驗員
                inspector = inspector & Cells(j, k) & Chr(13)
        
                Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                Range("F" & i) = inspector
            End If
            '----------檢驗員----------
            

            '----------工單數----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "QR Code_工單數") <> 0 Then   '判斷是否等於工單數
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                Range("G" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------工單數----------
            
            
            '----------製令單號----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "QR Code_製令工單") <> 0 Then   '判斷是否等於製令單號
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                Range("H" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------製令單號----------
            
            
            '----------客戶----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "客戶") <> 0 Then   '判斷是否等於客戶
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                Range("J" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------客戶----------
            
            
            '----------機種----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "QR Code_料號") <> 0 Then   '判斷是否等於料號
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                Range("K" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------機種----------
            

            '----------品名----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "品名") <> 0 Then   '判斷是否等於品名
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                Range("L" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------品名----------
            
            '----------SOP----------
            Workbooks(ActWb).Worksheets(1).Activate
    
            If InStr(1, Cells(1, k), "作業規範_SOP") <> 0 Then   '判斷是否等於SOP
                If InStr(1, Cells(j, k), "█") = 1 Then
                    Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                    Range("M" & i) = "V"
                Else
                    Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                    Range("M" & i) = "X"
                End If
            End If
            '----------SOP----------
            
            '----------SIP----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "作業規範_SIP") <> 0 Then   '判斷是否等於SIP
                If InStr(1, Range("J" & j), "█") = 1 Then
                    Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                    Range("N" & i) = "V"
                Else
                    Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                    Range("N" & i) = "X"
                End If
            End If
            '----------SIP----------
            
            '----------樣品----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "作業規範_樣品") <> 0 Then   '判斷是否等於樣品
                If InStr(1, Range("K" & j), "█") = 1 Then
                    Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                    Range("O" & i) = "V"
                Else
                    Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                    Range("O" & i) = "X"
                End If
            End If
            '----------樣品----------
                        
            
            '----------製造數----------
            Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
    
            If Range("D" & i) = "IPQC" Then    '判斷是IPQC，抓取半成品檢查數
                Workbooks(ActWb).Worksheets(1).Activate
                
                If InStr(1, Cells(1, k), "半成品檢查數") <> 0 Then   '判斷是否等於半成品檢查數
                
                    Dim Semi_Finished_Rroduct As Long
                    
                    Semi_Finished_Rroduct = Semi_Finished_Rroduct + Cells(j, k)

                    Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                    Range("R" & i) = Semi_Finished_Rroduct
                End If
                
            ElseIf Range("D" & i) = "首件" Then '判斷是首件，則為1
                Range("R" & i) = 1
            End If
            '----------製造數----------
            
        Next k
        
        inspector = ""
        Semi_Finished_Rroduct = 0
        
        '----------FQC----------
            Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
        
            If Range("D" & i) = "IPQC" Then    '判斷是 IPQC 且入庫數不為""，複製一筆 IPQC 改為 FQC
                Workbooks(ActWb).Worksheets(1).Activate
                
                For k = 1 To Workbooks(ActWb).Worksheets(1).Cells(1, Columns.Count).End(xlToLeft).Column '最後一欄
                    If InStr(1, Cells(1, k), "入庫數") <> 0 Then   '判斷是否等於入庫數
                        
                        If Cells(j, k) <> "" Then
                            Semi_Finished_Rroduct = Cells(j, k)
                            
                            Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate
                            Range("D" & i & ":AD" & i).Select
                            
                            MsgBox "select = " & Selection.Address
                            
                            Selection.Copy
                            Range("D" & i & ":AD" & i).Offset(1, 0).Select
                            
                            MsgBox "select = " & Selection.Address
                            
                            Selection.PasteSpecial xlPasteValues
                            
                            Range("D" & i).Offset(1, 0) = "FQC"
                            Range("R" & i).Offset(1, 0) = Semi_Finished_Rroduct
        
                        End If
                    End If
                Next k
            End If
            
            Semi_Finished_Rroduct = 0
            '----------FQC----------
        
    Next j
            
End Sub

