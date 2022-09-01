Sub GeneralTable_成型首件_0830前()

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name


'-----------------匯出資料總表整理-----------------
Range("A:F, H:H, L:N, EU:EW, FF:FF").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False


'Workbooks(ActWb).Worksheets(1).Activate

Application.CutCopyMode = False

'-----------------匯出資料總表整理-----------------

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

Range("L1") = "綜合判定"
Range("M1") = "檢驗員"
Range("N1") = "不良原因"
Range("O1") = "不良現象"
Range("P1") = "不良處理方式"

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
Range("S1") = "外觀_抽驗數"
Range("S2").Select
ActiveCell.Formula = "=IF(AND(H2>=2, H2<=544), 32, IF(AND(H2>=545, H2<=960), 40,  IF(AND(H2>=961, H2<=1632), 48,  IF(AND(H2>=1633, H2<=3072), 64,  IF(H2>=3073, 80, 1)))))"
Selection.AutoFill Destination:=Range("S2:S" & lrow)
'-----------------總表 外觀_抽驗數-----------------


'-----------------總表 VIP_抽驗數-----------------
Range("T1") = "抽驗數"
Range("T2").Select
ActiveCell.Formula = "=IF(AND(H2>=2, H2<=170), 5, IF(AND(H2>=171, H2<=288), 6,  IF(AND(H2>=289, H2<=544), 8,  IF(AND(H2>=545, H2<=960), 10,  IF(H2>=961, 12, 1)))))"
Selection.AutoFill Destination:=Range("T2:T" & lrow)
'-----------------總表 VIP_抽驗數-----------------


'-----------------總表 抽驗數_外觀+VIP-----------------
Range("U1") = "抽驗數_外觀+VIP"
Range("U2").Select
ActiveCell.Formula = "=S2+T2"
Selection.AutoFill Destination:=Range("U2:U" & lrow)
'-----------------總表 抽驗數_外觀+VIP-----------------


'-----------------總表 不良數-----------------
Range("V1") = "不良數"
Range("V2") = 0
Range("V2").Select
Selection.AutoFill Destination:=Range("V2:V" & lrow)
'-----------------總表 不良數-----------------


'-----------------總表 不良率-----------------
Range("W1") = "不良率"
Range("W2").Formula = "=""-"""
Range("W2").Select
Selection.AutoFill Destination:=Range("W2:W" & lrow)
'-----------------總表 不良率-----------------



'-----------------總表 批不良率-----------------
Range("X1") = "批不良率"
Range("X2").Formula = "-"
Range("X2").Select
Selection.AutoFill Destination:=Range("X2:X" & lrow)
'-----------------總表 批不良率-----------------



'-----------------總表 判定-----------------
Range("Y1") = "判定"
Range("Y2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IFERROR(IF(FIND(""可生產"",N2), ""合格"", ""不合格""),"""")"
Range("Y2").Select
Selection.AutoFill Destination:=Range("Y2:Y" & lrow)
'-----------------總表 判定-----------------



'-----------------總表 不良1原因-----------------
Range("Z1") = "不良1原因"
Range("Z2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(P2 = """","""", P2 & ""，"" & Q2 & ""，"" & R2)"
Range("Z2").Select
Selection.AutoFill Destination:=Range("Z2:Z" & lrow)
'-----------------總表 不良1原因-----------------


'-----------------總表 NG數資料-----------------
'For k = 2 To 5000
'
'    If Range("K" & k) = "NG" Then
'
'        If Range("A" & k) = Range("A" & k).Offset(-1, 0) And Range("C" & k) = Range("C" & k).Offset(-1, 0) Then
'            k = k + 1
'        Else
'            For m = 1 To Range("V" & k)
'
'                Range("A" & k & ":V" & k).Select
'                Selection.Copy
'
'                Range("A" & k & ":V" & k).Offset(1, 0).Select
'                Selection.Insert Shift:=xlDown
'            Next m
'
'            Range("K" & k) = "OK"
'            Range("S" & k) = 0
'        End If
'    End If
'Next
'-----------------總表 NG數資料-----------------



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

ActiveSheet.Range("O2", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select
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
ActiveSheet.Range("U2", ActiveSheet.Range("U" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗數外觀+VIP-----------------



'-----------------不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("V2", ActiveSheet.Range("V" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良數-----------------



'-----------------不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("W2", ActiveSheet.Range("W" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良率-----------------


'-----------------判定-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("Y2", ActiveSheet.Range("Y" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------


'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("X2", ActiveSheet.Range("X" & ActiveSheet.Rows.Count).End(xlUp)).Select
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

ActiveSheet.Range("Z2", ActiveSheet.Range("Z" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良1原因-----------------

End Sub

