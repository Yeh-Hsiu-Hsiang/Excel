Sub 成型IPQC_QRCode_0830前()

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

Range("A:F, N:P, W:X, AW:AW, CR:CT, DG:DI, DW:DY, GD:GD, GV:GW, IN:IO, IS:IX").Select
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
ActiveCell.FormulaR1C1 = "IPQC"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & lrow)
'-----------------總表 項目-----------------


'-----------------總表 機台-----------------
Columns("J:J").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("J1") = "機台"
Range("J2").Select
ActiveCell.Formula = "=ASC(I2)"
Selection.AutoFill Destination:=Range("J2:J" & lrow)
'-----------------總表 機台-----------------



'-----------------總表 抽驗數_外觀+VIP-----------------
Columns("AB:AB").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AB1") = "抽驗數_外觀+VIP"
Range("AB2").Select
ActiveCell.Formula = "=Z2+AA2"
Selection.AutoFill Destination:=Range("AB2:AB" & lrow)
'-----------------總表 抽驗數_外觀+VIP-----------------



'-----------------總表 不良數-----------------
Range("AK1") = "不良數總計"
Range("AK2").Select
ActiveCell.Formula = "=IF(AND(AH2="""", AI2="""", AJ2=""""), 0, AH2+AI2+AJ2)"
Selection.AutoFill Destination:=Range("AK2:AK" & lrow)
'-----------------總表 不良數-----------------


'-----------------總表 不良率-----------------
Range("AL1") = "不良率"
Range("AL2").Select
ActiveCell.Formula = "=IFERROR(AK2/AB2, 0)"
Selection.AutoFill Destination:=Range("AL2:AL" & lrow)
'-----------------總表 不良率-----------------


'-----------------總表 判定-----------------
Range("AM1") = "判定"
Range("AM2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(AK2=0, ""合格"", ""不合格"")"
Range("AM2").Select
Selection.AutoFill Destination:=Range("AM2:AM" & lrow)
'-----------------總表 判定-----------------



'-----------------總表 批不良率-----------------
Range("AN1") = "批不良率"
Range("AN2").Select
ActiveCell.Formula = "=IFERROR(AK2/H2, 0)"
Selection.AutoFill Destination:=Range("AN2:AN" & lrow)
'-----------------總表 批不良率-----------------



'-----------------總表 技術員-----------------
Range("AO1") = "技術員"
Range("AO2").Select
ActiveCell.Formula = "=IF(AND(L2="""",N2=""""),"""", L2 & "" "" & N2)"
Selection.AutoFill Destination:=Range("AO2:AO" & lrow)
'-----------------總表 技術員-----------------


'-----------------總表 不良1原因-----------------
Range("AP1") = "不良1原因"
Range("AP2").Select
ActiveCell.Formula = "=IF(P2 = """","""", P2 & ""，"" & Q2 & ""，"" & R2)"
Range("AP2").Select
Selection.AutoFill Destination:=Range("AP2:AP" & lrow)
'-----------------總表 不良1原因-----------------


'-----------------總表 不良2原因-----------------
Range("AQ1") = "不良2原因"
Range("AQ2").Select
ActiveCell.Formula = "=IF(S2 = """","""", S2 & ""，"" & T2 & ""，"" & U2)"
Range("AQ2").Select
Selection.AutoFill Destination:=Range("AQ2:AQ" & lrow)
'-----------------總表 不良2原因-----------------


'-----------------總表 不良3原因-----------------
Range("AR1") = "不良3原因"
Range("AR2").Select
ActiveCell.Formula = "=IF(V2 = """","""", V2 & ""，"" & W2 & ""，"" & X2)"
Range("AR2").Select
Selection.AutoFill Destination:=Range("AR2:AR" & lrow)
'-----------------總表 不良3原因-----------------


'-----------------總表 重工不良率-----------------
Range("AS1") = "重工不良率"
Range("AS2").Select
ActiveCell.Formula = "=IFERROR(AF2/AE2, 0)"
Range("AS2").Select
Selection.AutoFill Destination:=Range("AS2:AS" & lrow)
'-----------------總表 重工不良率-----------------



'-----------------總表 重工資訊-----------------
Range("AT1") = "重工資訊"
Range("AT2").Select
ActiveCell.Formula = "=IF(AE2="""","""",""重工數量 = "" & AE2)"
Range("AT2").Select
Selection.AutoFill Destination:=Range("AT2:AT" & lrow)
'-----------------總表 重工資訊-----------------


'-----------------總表 NG數-----------------
Range("AU1") = "NG數"
Range("AU2").Select
ActiveCell.Formula = "=IF(AK2>0, 1, 0)"
Range("AU2").Select
Selection.AutoFill Destination:=Range("AU2:AU" & lrow)
'-----------------總表 NG數-----------------


'-----------------總表 NG數資料-----------------
For k = 2 To 5000

    If Range("AM" & k) = "不合格" Then
        If Range("B" & k) = Range("B" & k).Offset(-1, 0) And _
            Range("E" & k) = Range("E" & k).Offset(-1, 0) And _
            Range("G" & k) = Range("G" & k).Offset(-1, 0) Then
        
            GoTo ContinueForLoop
        Else
            For m = 1 To Range("AU" & k)
                Range("A" & k & ":AU" & k).Select
                Selection.Copy

                Range("A" & k & ":AU" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next m

            Range("AK" & k) = 0
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
ActiveSheet.Range("D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("C" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------客戶-----------------



'-----------------製令單號-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("D" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令單號-----------------



'-----------------檢驗員A-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AC2", ActiveSheet.Range("AC" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員A-----------------



'-----------------檢驗員B-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AD2", ActiveSheet.Range("AD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("G" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員B-----------------


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
ActiveSheet.Range("J2", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).Select
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
ActiveSheet.Range("AB2", ActiveSheet.Range("AB" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗數外觀+VIP-----------------


'-----------------不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AK2", ActiveSheet.Range("AK" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良數-----------------


'-----------------不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AL2", ActiveSheet.Range("AL" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良率-----------------



'-----------------判定-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AM2", ActiveSheet.Range("AM" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------



'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AN2", ActiveSheet.Range("AN" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------批不良率-----------------



'-----------------技術員-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AO2", ActiveSheet.Range("AO" & ActiveSheet.Rows.Count).End(xlUp)).Select
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

ActiveSheet.Range("AP2", ActiveSheet.Range("AP" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良1原因-----------------


'-----------------不良2原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AQ2", ActiveSheet.Range("AQ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良2原因-----------------


'-----------------不良3原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AR2", ActiveSheet.Range("AR" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Z" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良3原因-----------------



'-----------------重工資訊(重工數)-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AT2", ActiveSheet.Range("AT" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工資訊(重工數)-----------------



'-----------------重工不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AF2", ActiveSheet.Range("AF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AD" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工不良數-----------------



'-----------------重工不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AS2", ActiveSheet.Range("AS" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AE" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工不良率-----------------

End Sub
