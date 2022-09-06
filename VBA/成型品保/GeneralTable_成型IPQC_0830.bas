Sub GeneralTable_成型IPQC_0830()

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

Range("A:G, N:P, Y:Z, AM:AM, BA:BA, BM:BO, CL:CL, CY:DB, DM:DM, DO:DR, EE:EH, ES:ES, FP:FP, GM:GM, HE:HF, IU:IU, IW:IX").Select
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
Columns("P:P").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("P1") = "IPQC判定_08~10時段"
Range("P2").Select
ActiveCell.Formula = "=""08~10(20~22)"""
Selection.AutoFill Destination:=Range("P2:P" & lrow)
'-----------------總表 IPQC判定_08~10-----------------



'-----------------總表 IPQC判定_10~12-----------------
Columns("S:S").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("S1") = "IPQC判定_10~12時段"
Range("S2").Select
ActiveCell.Formula = "=""10~12(22~24)"""
Selection.AutoFill Destination:=Range("S2:S" & lrow)
'-----------------總表 IPQC判定_10~12-----------------


'-----------------總表 IPQC判定_12~14-----------------
Columns("W:W").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("W1") = "IPQC判定_12~14時段"
Range("W2").Select
ActiveCell.Formula = "=""12~14(24~02)"""
Selection.AutoFill Destination:=Range("W2:W" & lrow)
'-----------------總表 IPQC判定_12~14-----------------


'-----------------總表 IPQC判定_14~16-----------------
Columns("AC:AC").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AC1") = "IPQC判定_14~16時段"
Range("AC2").Select
ActiveCell.Formula = "=""14~16(02~04)"""
Selection.AutoFill Destination:=Range("AC2:AC" & lrow)
'-----------------總表 IPQC判定_14~16-----------------



'-----------------總表 IPQC判定_16~18-----------------
Columns("AM:AM").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AM1") = "IPQC判定_16~18時段"
Range("AM2").Select
ActiveCell.Formula = "=""16~18(04~06)"""
Selection.AutoFill Destination:=Range("AM2:AM" & lrow)
'-----------------總表 IPQC判定_16~18-----------------


'-----------------總表 IPQC判定_18~20-----------------
Columns("AO:AO").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AO1") = "IPQC判定_18~20時段"
Range("AO2").Select
ActiveCell.Formula = "=""18~20(06~08)"""
Selection.AutoFill Destination:=Range("AO2:AO" & lrow)
'-----------------總表 IPQC判定_18~20-----------------


'-----------------總表 巡檢時段-----------------
Range("AV1") = "巡檢時段"
Range("AV2").Select
ActiveCell.Formula = "=IF(O2="""",IF(R2="""",IF(V2="""",IF(AB2="""",IF(AL2="""",IF(AN2="""","""",AO2),IF(AN2="""",AM2,AM2&"";""&AO2)),IF(AL2="""",IF(AN2="""",AC2,AC2&"";""&AO2),IF(AN2="""",AC2&"";""&AM2,AC2&"";""&AM2&"";""&AO2))),IF(AB2="""",IF(AL2="""",IF(AN2="""",W2,W2&"";""&AO2),IF(AN2="""",W2&"";""&AM2,W2&"";""&AM2&"";""&AO2)),IF(AL2="""",IF(AN2="""",W2&"";""&AC2,W2&"";""&AC2&"";""&AO2),IF(AN2="""",W2&"";""&AC2&"";""&AM2,W2&"";""&AC2&"";""&AM2&"";""&AO2)))),IF(V2="""",IF(AB2="""",IF(AL2="""",IF(AN2="""",S2,S2&"";""&AO2),IF(AN2="""",S2&"";""&AM2,S2&"";""&AM2&"";""&AO2))," & _
                     "IF(AL2="""",IF(AN2="""",S2&"";""&AC2,S2&"";""&AC2&"";""&AO2),IF(AN2="""",S2&"";""&AC2&"";""&AM2,S2&"";""&AC2&"";""&AM2&"";""&AO2))),IF(AB2="""",IF(AL2="""",IF(AN2="""",S2&"";""&W2,S2&"";""&W2&"";""&AO2),IF(AN2="""",S2&"";""&W2&"";""&AM2,S2&"";""&W2&"";""&AM2&"";""&AO2)),IF(AL2="""",IF(AN2="""",S2&"";""&W2&"";""&AC2,S2&"";""&W2&"";""&AC2&"";""&AO2),IF(AN2="""",S2&"";""&W2&"";""&AC2&"";""&AM2,S2&"";""&W2&"";""&AC2&"";""&AM2&"";""&AO2))))),IF(R2="""",IF(V2="""",IF(AB2="""",IF(AL2="""",IF(AN2="""",P2,P2&"";""&AO2),IF(AN2="""",P2&"";""&AM2,O2&"";""&AM2&"";""&AO2))," & _
                     "IF(AL2="""",IF(AN2="""",P2&"";""&AC2,P2&"";""&AC2&"";""&AO2),IF(AN2="""",P2&"";""&AC2&"";""&AM2,P2 &"";""&AC2&"";""&AM2&"";""&AO2))),IF(AB2="""",IF(AL2="""",IF(AN2="""",P2&"";""&W2,P2&"";""&W2&"";""&AO2),IF(AN2="""",P2&"";""&W2&"";""&AM2,P2&"";""&W2&"";""&AM2&"";""&AO2)),IF(AL2="""",IF(AN2="""",P2&"";""&W2&"";""&AC2,P2&"";""&W2&"";""&AC2&"";""&AO2),IF(AN2="""",P2&"";""&W2&"";""&AC2&"";""&AM2,P2&"";""&W2&"";""&AC2&"";""&AM2&"";""&AO2)))),IF(V2="""",IF(AB2="""",IF(AL2="""",IF(AN2="""",P2&"";""&S2,P2&"";""&S2&"";""&AO2),IF(AN2="""",P2&"";""&S2&"";""&AM2,P2&"";""&S2&"";""&AM2&"";""&AO2))," & _
                     "IF(AL2="""",IF(AN2="""",P2&"";""&S2&"";""&AC2,S2&"";""&AC2&"";""&AO2),IF(AN2="""",P2&"";""&S2&"";""&AC2&"";""&AM2,P2&"";""&S2&"";""&AC2&"";""&AM2&"";""&AO2))),IF(AB2="""",IF(AL2="""",IF(AN2="""",P2&"";""&S2&"";""&W2,P2&"";""&S2&"";""&W2&"";""&AO2),IF(AN2="""",P2&"";""&S2&"";""&W2&"";""&AM2,P2&"";""&S2&"";""&W2&"";""&AM2&"";""&AO2)),IF(AL2="""",IF(AN2="""",P2&"";""&S2&"";""&W2&"";""&AC2,P2&"";""&S2&"";""&W2&"";""&AC2&"";""&AO2),IF(AN2="""",P2&"";""&S2&"";""&W2&"";""&AC2&"";""&AM2,P2&"";""&S2&"";""&W2&"";""&AC2&"";""&AM2&"";""&AO2))))))"
Selection.AutoFill Destination:=Range("AV2:AV" & lrow)
'-----------------總表 巡檢時段-----------------


'-----------------總表 巡檢次數-----------------
Columns("AW:AW").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AW1") = "巡檢次數"
Range("AW2").Select
ActiveCell.Formula = "=COUNTA(O2,R2,V2,AB2,AL2,AN2)"
Selection.AutoFill Destination:=Range("AW2:AW" & lrow)
'-----------------總表 巡檢次數-----------------


'-----------------總表 機台-----------------
Columns("K:K").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("K1") = "機台"
Range("K2").Select
ActiveCell.Formula = "=ASC(J2)"
Selection.AutoFill Destination:=Range("K2:K" & lrow)
'-----------------總表 機台-----------------


'-----------------總表 抽驗數_外觀-----------------
Range("AR2").Select
ActiveCell.Formula = "=IF(AND(AT2>=2, AT2<=544), 32, IF(AND(AT2>=545, AT2<=960), 40,  IF(AND(AT2>=961, AT2<=1632), 48,  IF(AND(AT2>=1633, AT2<=3072), 64,  IF(AT2>=3073, 80, 1)))))"
Selection.AutoFill Destination:=Range("AR2:AR" & lrow)
'-----------------總表 抽驗數_外觀-----------------



'-----------------總表 抽驗數_VIP-----------------
Range("AS2").Select
ActiveCell.Formula = "=IF(AND(AT2>=2, AT2<=170), 5, IF(AND(AT2>=171, AT2<=288), 6,  IF(AND(AT2>=289, AT2<=544), 8,  IF(AND(AT2>=545, AT2<=960), 10,  IF(AT2>=961, 12, 1)))))"
Selection.AutoFill Destination:=Range("AS2:AS" & lrow)
'-----------------總表 抽驗數_VIP-----------------




'-----------------總表 抽驗數_外觀+VIP-----------------
Columns("AT:AT").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AT1") = "抽驗數_外觀+VIP"
Range("AT2").Select
ActiveCell.Formula = "=AR2+AS2"
Selection.AutoFill Destination:=Range("AT2:AT" & lrow)
'-----------------總表 抽驗數_外觀+VIP-----------------



'-----------------總表 不良數-----------------
Columns("AM:AM").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AM1") = "不良數總計"
Range("AM2").Select
ActiveCell.Formula = "=IF(AND(AB2="""", AH2="""", AL2=""""), 0, AB2+AH2+AL2)"
Selection.AutoFill Destination:=Range("AM2:AM" & lrow)
'-----------------總表 不良數-----------------


'-----------------總表 不良率-----------------
Range("BA1") = "不良率"
Range("BA2").Select
ActiveCell.Formula = "=IFERROR(AM2/AU2, 0)"
Selection.AutoFill Destination:=Range("BA2:BA" & lrow)
'-----------------總表 不良率-----------------



'-----------------總表 判定-----------------
Range("BB1") = "判定"
Range("BB2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(AM2=0, ""合格"", ""不合格"")"
Range("BB2").Select
Selection.AutoFill Destination:=Range("BB2:BB" & lrow)
'-----------------總表 判定-----------------


'-----------------總表 批不良率-----------------
Range("BC1") = "批不良率"
Range("BC2").Select
ActiveCell.Formula = "=IFERROR(AM2/AV2, 0)"
Selection.AutoFill Destination:=Range("BC2:BC" & lrow)
'-----------------總表 批不良率-----------------


'-----------------總表 技術員-----------------
Range("BD1") = "技術員"
Range("BD2").Select
ActiveCell.Formula = "=IF(AND(M2="""",O2=""""),"""", M2 & "" "" & O2)"
Selection.AutoFill Destination:=Range("BD2:BD" & lrow)
'-----------------總表 技術員-----------------


'-----------------總表 不良1原因-----------------
Range("BE1") = "不良1原因"
Range("BE2").Select
ActiveCell.Formula = "=IF(Y2 = """","""", Y2 & ""，"" & Z2 & ""，"" & AA2)"
Range("BE2").Select
Selection.AutoFill Destination:=Range("BE2:BE" & lrow)
'-----------------總表 不良1原因-----------------


'-----------------總表 不良2原因-----------------
Range("BF1") = "不良2原因"
Range("BF2").Select
ActiveCell.Formula = "=IF(AE2 = """","""", AE2 & ""，"" & AF2 & ""，"" & AG2)"
Range("BF2").Select
Selection.AutoFill Destination:=Range("BF2:BF" & lrow)
'-----------------總表 不良2原因-----------------


'-----------------總表 不良3原因-----------------
Range("BG1") = "不良3原因"
Range("BG2").Select
ActiveCell.Formula = "=IF(AI2 = """","""", AI2 & ""，"" & AJ2 & ""，"" & AK2)"
Range("BG2").Select
Selection.AutoFill Destination:=Range("BG2:BG" & lrow)
'-----------------總表 不良3原因-----------------


'-----------------總表 重工不良率-----------------
Range("BH1") = "重工不良率"
Range("BH2").Select
ActiveCell.Formula = "=IFERROR(V2/U2, 0)"
Range("BH2").Select
Selection.AutoFill Destination:=Range("BH2:BH" & lrow)
'-----------------總表 重工不良率-----------------


'-----------------總表 重工資訊-----------------
Range("BI1") = "重工資訊"
Range("BI2").Select
ActiveCell.Formula = "=IF(U2="""","""",""重工數量 = "" & U2)"
Range("BI2").Select
Selection.AutoFill Destination:=Range("BI2:BI" & lrow)
'-----------------總表 重工資訊-----------------


'-----------------總表 NG數-----------------
Range("BJ1") = "NG數"
Range("BJ2").Select
ActiveCell.Formula = "=IF(AM2>0, 1, 0)"
Range("BJ2").Select
Selection.AutoFill Destination:=Range("BJ2:BJ" & lrow)
'-----------------總表 NG數-----------------


'-----------------總表 NG數資料-----------------
For k = 2 To 5000

    If Range("BB" & k) = "不合格" Then
        If Range("C" & k) = Range("C" & k).Offset(-1, 0) And _
            Range("F" & k) = Range("F" & k).Offset(-1, 0) And _
            Range("H" & k) = Range("H" & k).Offset(-1, 0) Then
        
            GoTo ContinueForLoop
        Else
            For M = 1 To Range("BJ" & k)
                Range("A" & k & ":BJ" & k).Select
                Selection.Copy

                Range("A" & k & ":BJ" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next M

            Range("AM" & k) = 0
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

ActiveSheet.Range("AW2", ActiveSheet.Range("AW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員A-----------------


'-----------------檢驗員B-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AX2", ActiveSheet.Range("AX" & ActiveSheet.Rows.Count).End(xlUp)).Select
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
ActiveSheet.Range("AY2", ActiveSheet.Range("AY" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("J" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------巡檢時段-----------------


'-----------------巡檢次數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AZ2", ActiveSheet.Range("AZ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("K" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------巡檢次數-----------------





'-----------------機台-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("L" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------機台-----------------


'-----------------生產數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AV2", ActiveSheet.Range("AV" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("M" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------生產數-----------------



'-----------------檢驗數外觀+VIP-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AU2", ActiveSheet.Range("AU" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗數外觀+VIP-----------------


'-----------------不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AM2", ActiveSheet.Range("AM" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良數-----------------


'-----------------不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良率-----------------


'-----------------判定-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("BB2", ActiveSheet.Range("BB" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------



'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("BC2", ActiveSheet.Range("BC" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------批不良率-----------------


'-----------------技術員-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BD2", ActiveSheet.Range("BD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------技術員-----------------


'-----------------作業員1-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("L2", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員1-----------------


'-----------------作業員2-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("N2", ActiveSheet.Range("N" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員2-----------------


'-----------------作業員3-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("R2", ActiveSheet.Range("R" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------作業員3-----------------


'-----------------不良1原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BE2", ActiveSheet.Range("BE" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良1原因-----------------



'-----------------不良2原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BF2", ActiveSheet.Range("BF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良2原因-----------------


'-----------------不良3原因-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BG2", ActiveSheet.Range("BG" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("Z" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良3原因-----------------


'-----------------重工資訊(重工數)-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BI2", ActiveSheet.Range("BI" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工資訊(重工數)-----------------



'-----------------重工不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("V2", ActiveSheet.Range("V" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AD" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工不良數-----------------


'-----------------重工不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BH2", ActiveSheet.Range("BH" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("AE" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------重工不良率-----------------

End Sub
