Sub GeneralTable_成型首件()

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name


'-----------------匯出資料總表整理-----------------
Range("A:F, H:H, K:N, EU:FA, FF:FF").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False


'-----------------匯出資料總表整理-----------------

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row


'-----------------總表 日期-----------------
Columns("B:B").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("B1") = "日期"
Range("B2").Select
ActiveCell.Formula = "=LEFT(A2, 4) & ""/"" & MID(A2, 5, 2) & ""/"" & RIGHT(A2, 2)"
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B" & lrow)
'-----------------總表 日期-----------------


'-----------------總表 項目-----------------
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("C1") = "項目"
Range("C2").Select
ActiveCell.FormulaR1C1 = "首件"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & lrow)
'-----------------總表 項目-----------------


'-----------------總表 外觀_抽驗數-----------------
Range("V1") = "外觀_抽驗數"
Range("V2").Select
ActiveCell.Formula = "=IF(AND(H2>=2, H2<=544), 32, IF(AND(H2>=545, H2<=960), 40,  IF(AND(H2>=961, H2<=1632), 48,  IF(AND(H2>=1633, H2<=3072), 64,  IF(H2>=3073, 80, 1)))))"
Selection.AutoFill Destination:=Range("V2:V" & lrow)
'-----------------總表 外觀_抽驗數-----------------


'-----------------總表 VIP_抽驗數-----------------
Range("W1") = "抽驗數"
Range("W2").Select
ActiveCell.Formula = "=IF(AND(H2>=2, H2<=170), 5, IF(AND(H2>=171, H2<=288), 6,  IF(AND(H2>=289, H2<=544), 8,  IF(AND(H2>=545, H2<=960), 10,  IF(H2>=961, 12, 1)))))"
Selection.AutoFill Destination:=Range("W2:W" & lrow)
'-----------------總表 VIP_抽驗數-----------------


'-----------------總表 抽驗數_外觀+VIP-----------------
Range("X1") = "抽驗數_外觀+VIP"
Range("X2").Select
ActiveCell.Formula = "=V2+W2"
Selection.AutoFill Destination:=Range("X2:X" & lrow)
'-----------------總表 抽驗數_外觀+VIP-----------------


'-----------------總表 不良數-----------------
Range("Y1") = "不良數"
Range("Y2").Select
ActiveCell.Formula = "=IF(T2>=2, (T2-1)*2, 0)"
Selection.AutoFill Destination:=Range("Y2:Y" & lrow)
'-----------------總表 不良數-----------------



'-----------------總表 不良率-----------------
Range("Z1") = "不良率"
Range("Z2").Select
ActiveCell.Formula = "=IFERROR(Y2/W2, 0)"
Selection.AutoFill Destination:=Range("Z2:Z" & lrow)
'-----------------總表 不良率-----------------



'-----------------總表 判定-----------------
Range("AA1") = "判定"
Range("AA2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(T2="""","""", IF(T2=1, ""合格"", ""不合格""))"
Range("AA2").Select
Selection.AutoFill Destination:=Range("AA2:AA" & lrow)
'-----------------總表 判定-----------------


'-----------------總表 批不良率-----------------
Range("AB1") = "批不良率"
Range("AB2").Select
ActiveCell.Formula = "=IFERROR(Y2/H2, 0)"
Selection.AutoFill Destination:=Range("AB2:AB" & lrow)
'-----------------總表 批不良率-----------------


'-----------------總表 不良1原因-----------------
Range("AC1") = "不良1原因"
Range("AC2").Select
ActiveCell.Formula = "=IF(N2 = """","""", N2 & ""，"" & O2 & ""，"" & P2)"
Range("AC2").Select
Selection.AutoFill Destination:=Range("AC2:AC" & lrow)
'-----------------總表 不良1原因-----------------


'-----------------總表 不良2原因-----------------
Range("AD1") = "不良2原因"
Range("AD2").Select
ActiveCell.Formula = "=IF(Q2 = """","""", Q2 & ""，"" & R2 & ""，"" & S2)"
Range("AD2").Select
Selection.AutoFill Destination:=Range("AD2:AD" & lrow)
'-----------------總表 不良2原因-----------------


'-----------------總表 NG數-----------------
Range("AE1") = "NG數"
Range("AE2").Select
ActiveCell.Formula = "=IF(T2="""", 0, IF(T2>=2, T2-1, 0))"
Range("AE2").Select
Selection.AutoFill Destination:=Range("AE2:AE" & lrow)
'-----------------總表 NG數-----------------


'-----------------總表 NG數資料-----------------
For k = 2 To 5000

    If Range("AA" & k) = "不合格" Then
        If Range("B" & k) = Range("B" & k).Offset(-1, 0) And Range("D" & k) = Range("D" & k).Offset(-1, 0) Then
            GoTo ContinueForLoop
        Else
            For m = 1 To Range("AE" & k)
                Range("A" & k & ":AE" & k).Select
                Selection.Copy

                Range("A" & k & ":AE" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next m

            Range("AA" & k) = "合格"
            Range("Y" & k) = 0
        End If
    End If
    
ContinueForLoop:
            Next k
'-----------------總表 NG數資料-----------------


Application.CutCopyMode = False


'-----------------匯出資料總表整理-----------------


'複製資料匯出總表 準備貼到品保 IPQC 總表
Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy


'-----------------C 欄 首件-----------------
Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

j = 6
Do While True
    If Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Cells(j, "A").Value = "" Then
        Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Cells(j, "A").Select
        Exit Do
    End If
    j = j + 1
Loop



ActiveSheet.Range("A" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------C 欄首件-----------------



'-----------------日期-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("B2", ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("B" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------日期-----------------


'-----------------客戶-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("C" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------客戶-----------------


'-----------------製令單號-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("D" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令單號-----------------


'-----------------班別-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("M2", ActiveSheet.Range("M" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("E" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------班別-----------------


'-----------------檢驗員-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("U2", ActiveSheet.Range("U" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員-----------------



'-----------------料號-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------料號-----------------


'-----------------品名-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("F2", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("I" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------品名-----------------




'-----------------機台-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("I2", ActiveSheet.Range("I" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("L" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------機台-----------------


'-----------------生產數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("M" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------生產數-----------------



'-----------------檢驗數外觀+VIP-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("X2", ActiveSheet.Range("X" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗數外觀+VIP-----------------



'-----------------不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("Y2", ActiveSheet.Range("Y" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良數-----------------



'-----------------不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("Z2", ActiveSheet.Range("Z" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良率-----------------


'-----------------判定-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AA2", ActiveSheet.Range("AA" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------


'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AB2", ActiveSheet.Range("AB" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------批不良率-----------------


'-----------------技術員-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("L2", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------技術員-----------------


'-----------------作業員1-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("J2", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員1-----------------


'-----------------作業員2-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員2-----------------


'-----------------不良1原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AC2", ActiveSheet.Range("AC" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良1原因-----------------


'-----------------不良2原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AD2", ActiveSheet.Range("AD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良2原因-----------------



'-----------------首件NG次數-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("T2", ActiveSheet.Range("T" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AB" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------首件NG次數-----------------


End Sub
