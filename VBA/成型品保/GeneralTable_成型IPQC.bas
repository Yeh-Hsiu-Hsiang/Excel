Sub GeneralTable_成型IPQC()

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

Range("A:F, N:P, V:W, AH:AH, AU:AU, BG:BI, CE:CE, CR:CU, DF:DF, DH:DK, DX:EA, EL:EL, FI:FI, GF:GF, GX:GY, IN:IN, IP:IQ").Select
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


'-----------------總表 IPQC判定_08~10-----------------
Columns("O:O").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("O1") = "IPQC判定_08~10時段"
Range("O2").Select
ActiveCell.Formula = "=""08~10(20~22)"""
Selection.AutoFill Destination:=Range("O2:O" & lrow)
'-----------------總表 IPQC判定_08~10-----------------


'-----------------總表 IPQC判定_10~12-----------------
Columns("R:R").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("R1") = "IPQC判定_10~12時段"
Range("R2").Select
ActiveCell.Formula = "=""10~12(22~24)"""
Selection.AutoFill Destination:=Range("R2:R" & lrow)
'-----------------總表 IPQC判定_10~12-----------------


'-----------------總表 IPQC判定_12~14-----------------
Columns("V:V").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("V1") = "IPQC判定_12~14時段"
Range("V2").Select
ActiveCell.Formula = "=""12~14(24~02)"""
Selection.AutoFill Destination:=Range("V2:V" & lrow)
'-----------------總表 IPQC判定_12~14-----------------


'-----------------總表 IPQC判定_14~16-----------------
Columns("AB:AB").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AB1") = "IPQC判定_14~16時段"
Range("AB2").Select
ActiveCell.Formula = "=""14~16(02~04)"""
Selection.AutoFill Destination:=Range("AB2:AB" & lrow)
'-----------------總表 IPQC判定_14~16-----------------



'-----------------總表 IPQC判定_16~18-----------------
Columns("AL:AL").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AL1") = "IPQC判定_16~18時段"
Range("AL2").Select
ActiveCell.Formula = "=""16~18(04~06)"""
Selection.AutoFill Destination:=Range("AL2:AL" & lrow)
'-----------------總表 IPQC判定_16~18-----------------


'-----------------總表 IPQC判定_18~20-----------------
Columns("AN:AN").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AN1") = "IPQC判定_18~20時段"
Range("AN2").Select
ActiveCell.Formula = "=""18~20(06~08)"""
Selection.AutoFill Destination:=Range("AN2:AN" & lrow)
'-----------------總表 IPQC判定_18~20-----------------



'-----------------總表 巡檢時段-----------------
Range("AU1") = "巡檢時段"
Range("AU2").Select
ActiveCell.Formula = "=IF(N2="""",IF(Q2="""",IF(U2="""",IF(AA2="""",IF(AK2="""",IF(AM2="""","""",AN2),IF(AM2="""",AL2,AL2&"";""&AN2)),IF(AK2="""",IF(AM2="""",AB2,AB2&"";""&AN2),IF(AM2="""",AB2&"";""&AL2,AB2&"";""&AL2&"";""&AN2))),IF(AA2="""",IF(AK2="""",IF(AM2="""",V2,V2&"";""&AN2),IF(AM2="""",V2&"";""&AL2,V2&"";""&AL2&"";""&AN2)),IF(AK2="""",IF(AM2="""",V2&"";""&AB2,V2&"";""&AB2&"";""&AN2),IF(AM2="""",V2&"";""&AB2&"";""&AL2,V2&"";""&AB2&"";""&AL2&"";""&AN2)))),IF(U2="""",IF(AA2="""",IF(AK2="""",IF(AM2="""",R2,R2&"";""&AN2),IF(AM2="""",R2&"";""&AL2,R2&"";""&AL2&"";""&AN2))," & _
                     "IF(AK2="""",IF(AM2="""",R2&"";""&AB2,R2&"";""&AB2&"";""&AN2),IF(AM2="""",R2&"";""&AB2&"";""&AL2,R2&"";""&AB2&"";""&AL2&"";""&AN2))),IF(AA2="""",IF(AK2="""",IF(AM2="""",R2&"";""&V2,R2&"";""&V2&"";""&AN2),IF(AM2="""",R2&"";""&V2&"";""&AL2,R2&"";""&V2&"";""&AL2&"";""&AN2)),IF(AK2="""",IF(AM2="""",R2&"";""&V2&"";""&AB2,R2&"";""&V2&"";""&AB2&"";""&AN2),IF(AM2="""",R2&"";""&V2&"";""&AB2&"";""&AL2,R2&"";""&V2&"";""&AB2&"";""&AL2&"";""&AN2))))),IF(Q2="""",IF(U2="""",IF(AA2="""",IF(AK2="""",IF(AM2="""",O2,O2&"";""&AN2), IF(AM2="""",O2&"";""&AL2,O2&"";""&AL2&"";""&AN2)),IF(AK2="""",IF(AM2="""",O2&"";""&AB2,O2&"";""&AB2&"";""&AN2)," & _
                     "IF(AM2="""",O2&"";""&AB2&"";""&AL2,O2&"";""&AB2&"";""&AL2&"";""&AN2))),IF(AA2="""",IF(AK2="""",IF(AM2="""",O2&"";""&V2,O2&"";""&V2&"";""&AN2),IF(AM2="""",O2&"";""&V2&"";""&AL2,O2&"";""&V2&"";""&AL2&"";""&AN2)),IF(AK2="""",IF(AM2="""",O2&"";""&V2&"";""&AB2,O2&"";""&V2&"";""&AB2&"";""&AN2),IF(AM2="""",O2&"";""&V2&"";""&AB2&"";""&AL2,O2&"";""&V2&"";""&AB2&"";""&AL2&"";""&AN2)))),IF(U2="""",IF(AA2="""",IF(AK2="""",IF(AM2="""",O2&"";""&R2,O2&"";""&R2&"";""&AN2),IF(AM2="""",O2&"";""&R2&"";""&AL2,O2&"";""&R2&"";""&AL2&"";""&AN2)),IF(AK2="""",IF(AM2="""",O2&"";""&R2&"";""&AB2,R2&"";""&AB2&"";""&AN2),IF(AM2="""",O2&"";""&R2&"";""&AB2&"";""&AL2,O2&"";""&R2&"";""&AB2&"";""&AL2&"";""&AN2)))," & _
                     "IF(AA2="""",IF(AK2="""",IF(AM2="""",O2&"";""&R2&"";""&V2,O2&"";""&R2&"";""&V2&"";""&AN2),IF(AM2="""",O2&"";""&R2&"";""&V2&"";""&AL2,O2&"";""&R2&"";""&V2&"";""&AL2&"";""&AN2)),IF(AK2="""",IF(AM2="""",O2&"";""&R2&"";""&V2&"";""&AB2,O2&"";""&R2&"";""&V2&"";""&AB2&"";""&AN2),IF(AM2="""",O2&"";""&R2&"";""&V2&"";""&AB2&"";""&AL2,O2&"";""&R2&"";""&V2&"";""&AB2&"";""&AL2&"";""&AN2))))))"
Selection.AutoFill Destination:=Range("AU2:AU" & lrow)
'-----------------總表 巡檢時段-----------------


'-----------------總表 巡檢次數-----------------
Columns("AV:AV").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AV1") = "巡檢次數"
Range("AV2").Select
ActiveCell.Formula = "=COUNTA(N2,Q2,U2,AA2,AK2,AM2)"
Selection.AutoFill Destination:=Range("AV2:AV" & lrow)
'-----------------總表 巡檢次數-----------------


'-----------------總表 抽驗數_外觀-----------------
Range("AP2").Select
ActiveCell.Formula = "=IF(AND(AR2>=2, AR2<=544), 32, IF(AND(AR2>=545, AR2<=960), 40,  IF(AND(AR2>=961, AR2<=1632), 48,  IF(AND(AR2>=1633, AR2<=3072), 64,  IF(AR2>=3073, 80, 1)))))"
Selection.AutoFill Destination:=Range("AP2:AP" & lrow)
'-----------------總表 抽驗數_外觀-----------------



'-----------------總表 抽驗數_VIP-----------------
Range("AQ2").Select
ActiveCell.Formula = "=IF(AND(AR2>=2, AR2<=170), 5, IF(AND(AR2>=171, AR2<=288), 6,  IF(AND(AR2>=289, AR2<=544), 8,  IF(AND(AR2>=545, AR2<=960), 10,  IF(AR2>=961, 12, 1)))))"
Selection.AutoFill Destination:=Range("AQ2:AQ" & lrow)
'-----------------總表 抽驗數_VIP-----------------



'-----------------總表 抽驗數_外觀+VIP-----------------
Columns("AR:AR").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AR1") = "抽驗數_外觀+VIP"
Range("AR2").Select
ActiveCell.Formula = "=AP2+AQ2"
Selection.AutoFill Destination:=Range("AR2:AR" & lrow)
'-----------------總表 抽驗數_外觀+VIP-----------------



'-----------------總表 不良數-----------------
Columns("AK:AK").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AK1") = "不良數總計"
Range("AK2").Select
ActiveCell.Formula = "=IF(AND(Z2="""", AF2="""", AJ2=""""), 0, Z2+AF2+AJ2)"
Selection.AutoFill Destination:=Range("AK2:AK" & lrow)
'-----------------總表 不良數-----------------



'-----------------總表 不良率-----------------
Range("AY1") = "不良率"
Range("AY2").Select
ActiveCell.Formula = "=IFERROR(AK2/AS2, 0)"
Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
'-----------------總表 不良率-----------------


'-----------------總表 判定-----------------
Range("AZ1") = "判定"
Range("AZ2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(AK2=0, ""合格"", ""不合格"")"
Range("AZ2").Select
Selection.AutoFill Destination:=Range("AZ2:AZ" & lrow)
'-----------------總表 判定-----------------


'-----------------總表 批不良率-----------------
Range("BA1") = "批不良率"
Range("BA2").Select
ActiveCell.Formula = "=IFERROR(AK2/AT2, 0)"
Selection.AutoFill Destination:=Range("BA2:BA" & lrow)
'-----------------總表 批不良率-----------------


'-----------------總表 技術員-----------------
Range("BB1") = "技術員"
Range("BB2").Select
ActiveCell.Formula = "=IF(AND(K2="""",M2=""""),"""", K2 & "" "" & M2)"
Selection.AutoFill Destination:=Range("BB2:BB" & lrow)
'-----------------總表 技術員-----------------



'-----------------總表 不良1原因-----------------
Range("BC1") = "不良1原因"
Range("BC2").Select
ActiveCell.Formula = "=IF(W2 = """","""", W2 & ""，"" & X2 & ""，"" & Y2)"
Range("BC2").Select
Selection.AutoFill Destination:=Range("BC2:BC" & lrow)
'-----------------總表 不良1原因-----------------


'-----------------總表 不良2原因-----------------
Range("BD1") = "不良2原因"
Range("BD2").Select
ActiveCell.Formula = "=IF(AC2 = """","""", AC2 & ""，"" & AD2 & ""，"" & AE2)"
Range("BD2").Select
Selection.AutoFill Destination:=Range("BD2:BD" & lrow)
'-----------------總表 不良2原因-----------------



'-----------------總表 不良3原因-----------------
Range("BE1") = "不良3原因"
Range("BE2").Select
ActiveCell.Formula = "=IF(AG2 = """","""", AG & ""，"" & AH2 & ""，"" & AI2)"
Range("BE2").Select
Selection.AutoFill Destination:=Range("BE2:BE" & lrow)
'-----------------總表 不良3原因-----------------


'-----------------總表 重工不良率-----------------
Range("BF1") = "重工不良率"
Range("BF2").Select
ActiveCell.Formula = "=IFERROR(T2/S2, 0)"
Range("BF2").Select
Selection.AutoFill Destination:=Range("BF2:BF" & lrow)
'-----------------總表 重工不良率-----------------


'-----------------總表 重工資訊-----------------
Range("BG1") = "重工資訊"
Range("BG2").Select
ActiveCell.Formula = "=IF(S2="""","""",""重工數量 = "" & S2)"
Range("BG2").Select
Selection.AutoFill Destination:=Range("BG2:BG" & lrow)
'-----------------總表 重工資訊-----------------



'-----------------總表 NG數-----------------
Range("BH1") = "NG數"
Range("BH2").Select
ActiveCell.Formula = "=IF(AK2>0, 1, 0)"
Range("BH2").Select
Selection.AutoFill Destination:=Range("BH2:BH" & lrow)
'-----------------總表 NG數-----------------



'-----------------總表 NG數資料-----------------
For k = 2 To 5000

    If Range("AZ" & k) = "不合格" Then
        If Range("C" & k) = Range("C" & k).Offset(-1, 0) And _
            Range("F" & k) = Range("F" & k).Offset(-1, 0) And _
            Range("H" & k) = Range("H" & k).Offset(-1, 0) Then

            GoTo ContinueForLoop
        Else

            For M = 1 To Range("BH" & k)
                Range("A" & k & ":BH" & k).Select
                Selection.Copy

                Range("A" & k & ":BH" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next M

            Range("AK" & k) = 0
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

ActiveSheet.Range("AU2", ActiveSheet.Range("AU" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員A-----------------


'-----------------檢驗員B-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AV2", ActiveSheet.Range("AV" & ActiveSheet.Rows.Count).End(xlUp)).Select
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



'-----------------巡檢時段-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AW2", ActiveSheet.Range("AW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("J" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------巡檢時段-----------------


'-----------------巡檢次數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AX2", ActiveSheet.Range("AX" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("K" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------巡檢次數-----------------



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
ActiveSheet.Range("AT2", ActiveSheet.Range("AT" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("M" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------生產數-----------------


'-----------------檢驗數外觀+VIP-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AS2", ActiveSheet.Range("AS" & ActiveSheet.Rows.Count).End(xlUp)).Select
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
ActiveSheet.Range("AY2", ActiveSheet.Range("AY" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良率-----------------


'-----------------判定-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AZ2", ActiveSheet.Range("AZ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------



'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------批不良率-----------------



'-----------------技術員-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BB2", ActiveSheet.Range("BB" & ActiveSheet.Rows.Count).End(xlUp)).Select
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

ActiveSheet.Range("L2", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員2-----------------



'-----------------作業員3-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("P2", ActiveSheet.Range("P" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員3-----------------


'-----------------不良1原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BC2", ActiveSheet.Range("BC" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良1原因-----------------


'-----------------不良2原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BD2", ActiveSheet.Range("BD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良2原因-----------------



'-----------------不良3原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BE2", ActiveSheet.Range("BE" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Z" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良3原因-----------------


'-----------------重工資訊(重工數)-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BG2", ActiveSheet.Range("BG" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工資訊(重工數)-----------------



'-----------------重工不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("T2", ActiveSheet.Range("T" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AD" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工不良數-----------------



'-----------------重工不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BF2", ActiveSheet.Range("BF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AE" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工不良率-----------------

Application.CutCopyMode = False

End Sub
