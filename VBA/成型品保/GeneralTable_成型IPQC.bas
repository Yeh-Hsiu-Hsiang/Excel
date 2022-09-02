Sub GeneralTable_成型IPQC()

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

Range("A:G, N:P, V:W, AU:AU, BH:BI, CR:CU, DH:DK, DX:EA, GF:GF, GX:GY, IP:IQ").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False


'-----------------匯出資料總表整理-----------------

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

'-----------------總表 日期-----------------
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("C1") = "日期"
Range("C2").Select
ActiveCell.Formula = "=LEFT(B2, 4) & ""/"" & MID(B2, 5, 2) & ""/"" & RIGHT(B2, 2)"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & lrow)
'-----------------總表 日期-----------------


'-----------------總表 項目-----------------
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("D1") = "項目"
Range("D2").Select
ActiveCell.FormulaR1C1 = "IPQC"
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D" & lrow)
'-----------------總表 項目-----------------


'-----------------總表 抽驗數_外觀+VIP-----------------
Columns("AG:AG").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AG1") = "抽驗數_外觀+VIP"
Range("AG2").Select
ActiveCell.Formula = "=AE2+AF2"
Selection.AutoFill Destination:=Range("AG2:AG" & lrow)
'-----------------總表 抽驗數_外觀+VIP-----------------



'-----------------總表 不良數-----------------
Columns("AD:AD").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AD1") = "不良數總計"
Range("AD2").Select
ActiveCell.Formula = "=IF(AND(U2="""", Y2="""", AC2=""""), 0, U2+Y2+AC2)"
Selection.AutoFill Destination:=Range("AD2:AD" & lrow)
'-----------------總表 不良數-----------------



'-----------------總表 不良率-----------------
Range("AK1") = "不良率"
Range("AK2").Select
ActiveCell.Formula = "=IFERROR(AD2/AH2, 0)"
Selection.AutoFill Destination:=Range("AK2:AK" & lrow)
'-----------------總表 不良率-----------------


'-----------------總表 判定-----------------
Range("AL1") = "判定"
Range("AL2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(AD2=0, ""合格"", ""不合格"")"
Range("AL2").Select
Selection.AutoFill Destination:=Range("AL2:AL" & lrow)
'-----------------總表 判定-----------------


'-----------------總表 批不良率-----------------
Range("AM1") = "批不良率"
Range("AM2").Select
ActiveCell.Formula = "=IFERROR(AD2/I2, 0)"
Selection.AutoFill Destination:=Range("AM2:AM" & lrow)
'-----------------總表 批不良率-----------------


'-----------------總表 技術員-----------------
Range("AN1") = "技術員"
Range("AN2").Select
ActiveCell.Formula = "=IF(AND(L2="""",N2=""""),"""", L2 & "" "" & N2)"
Selection.AutoFill Destination:=Range("AN2:AN" & lrow)
'-----------------總表 技術員-----------------



'-----------------總表 不良1原因-----------------
Range("AO1") = "不良1原因"
Range("AO2").Select
ActiveCell.Formula = "=IF(R2 = """","""", R2 & ""，"" & S2 & ""，"" & T2)"
Range("AO2").Select
Selection.AutoFill Destination:=Range("AO2:AO" & lrow)
'-----------------總表 不良1原因-----------------


'-----------------總表 不良2原因-----------------
Range("AP1") = "不良2原因"
Range("AP2").Select
ActiveCell.Formula = "=IF(V2 = """","""", V2 & ""，"" & W2 & ""，"" & X2)"
Range("AP2").Select
Selection.AutoFill Destination:=Range("AP2:AP" & lrow)
'-----------------總表 不良2原因-----------------



'-----------------總表 不良3原因-----------------
Range("AQ1") = "不良3原因"
Range("AQ2").Select
ActiveCell.Formula = "=IF(Z2 = """","""", Z & ""，"" & AA2 & ""，"" & AB2)"
Range("AQ2").Select
Selection.AutoFill Destination:=Range("AQ2:AQ" & lrow)
'-----------------總表 不良3原因-----------------


'-----------------總表 重工不良率-----------------
Range("AR1") = "重工不良率"
Range("AR2").Select
ActiveCell.Formula = "=IFERROR(Q2/P2, 0)"
Range("AR2").Select
Selection.AutoFill Destination:=Range("AR2:AR" & lrow)
'-----------------總表 重工不良率-----------------


'-----------------總表 重工資訊-----------------
Range("AS1") = "重工資訊"
Range("AS2").Select
ActiveCell.Formula = "=IF(P2="""","""",""重工數量 = "" & P2)"
Range("AS2").Select
Selection.AutoFill Destination:=Range("AS2:AS" & lrow)
'-----------------總表 重工資訊-----------------



'-----------------總表 NG數-----------------
Range("AT1") = "NG數"
Range("AT2").Select
ActiveCell.Formula = "=IF(AD2>0, 1, 0)"
Range("AT2").Select
Selection.AutoFill Destination:=Range("AT2:AT" & lrow)
'-----------------總表 NG數-----------------



'-----------------總表 NG數資料-----------------
For k = 2 To 5000

    If Range("AL" & k) = "不合格" Then
        If Range("C" & k) = Range("C" & k).Offset(-1, 0) And _
            Range("F" & k) = Range("F" & k).Offset(-1, 0) And _
            Range("H" & k) = Range("H" & k).Offset(-1, 0) Then

            GoTo ContinueForLoop
        Else
            
            For m = 1 To Range("AT" & k)
                Range("A" & k & ":AT" & k).Select
                Selection.Copy

                Range("A" & k & ":AT" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next m

            Range("AD" & k) = 0
        End If
    End If

ContinueForLoop:
            Next k
'-----------------總表 NG數資料-----------------


Application.CutCopyMode = False


'-----------------匯出資料總表整理-----------------


'複製資料匯出總表 準備貼到品保 IPQC 總表
Range("D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp)).Select
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

ActiveSheet.Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("B" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------日期-----------------


'-----------------客戶-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("C" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------客戶-----------------



'-----------------製令單號-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("D" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令單號-----------------



'-----------------班別-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("A2", ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("E" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------班別-----------------



'-----------------檢驗員A-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AI2", ActiveSheet.Range("AI" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員A-----------------


'-----------------檢驗員B-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AJ2", ActiveSheet.Range("AJ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("G" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員B-----------------



'-----------------料號-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("F2", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------料號-----------------


'-----------------品名-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("I" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------品名-----------------


'-----------------機台-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("J2", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("L" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------機台-----------------



'-----------------生產數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AE2", ActiveSheet.Range("AE" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("M" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------生產數-----------------


'-----------------檢驗數外觀+VIP-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AH2", ActiveSheet.Range("AH" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗數外觀+VIP-----------------



'-----------------不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AD2", ActiveSheet.Range("AD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良數-----------------



'-----------------不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AK2", ActiveSheet.Range("AK" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良率-----------------


'-----------------判定-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AL2", ActiveSheet.Range("AL" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------



'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AM2", ActiveSheet.Range("AM" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------批不良率-----------------



'-----------------技術員-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AN2", ActiveSheet.Range("AN" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------技術員-----------------



'-----------------作業員1-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員1-----------------



'-----------------作業員2-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("M2", ActiveSheet.Range("M" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員2-----------------



'-----------------作業員3-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("O2", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員3-----------------


'-----------------不良1原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AO2", ActiveSheet.Range("AO" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良1原因-----------------


'-----------------不良2原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AP2", ActiveSheet.Range("AP" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良2原因-----------------



'-----------------不良3原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AQ2", ActiveSheet.Range("AQ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Z" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良3原因-----------------


'-----------------重工資訊(重工數)-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AS2", ActiveSheet.Range("AS" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工資訊(重工數)-----------------



'-----------------重工不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("Q2", ActiveSheet.Range("Q" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AD" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工不良數-----------------



'-----------------重工不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AR2", ActiveSheet.Range("AR" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AE" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工不良率-----------------


End Sub
