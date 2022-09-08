Sub GeneralTable_成型FQC_0906()


Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

Range("A:H, O:Q, W:X, AI:AI, AV:AV, BH:BJ, CF:CF, CS:CV, DG:DG, DI:DL, DY:EB, EM:EM, FJ:FJ, GG:GG, GY:GZ, IO:IO, IQ:IR").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(2)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False

Columns("A:A").Select
Selection.NumberFormatLocal = "yyyy/mm/dd hh:mm"
Application.WindowState = xlNormal


'-----------------匯出資料總表整理-----------------

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

'-----------------總表 日期-----------------
Columns("D:D").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("D1") = "日期"
Range("D2").Select
ActiveCell.Formula = "=LEFT(C2, 4) & ""/"" & MID(C2, 5, 2) & ""/"" & RIGHT(C2, 2)"
Range("D2").Select
Selection.AutoFill Destination:=Range("D2:D" & lrow)
'-----------------總表 日期-----------------


'-----------------總表 項目-----------------
Columns("B:B").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("B1") = "項目"
Range("B2").Select
ActiveCell.Formula = "=IF(AK2<>"""", ""FQC"", """")"
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B" & lrow)
'-----------------總表 項目-----------------

For i = 2 To Range("A65536").End(xlUp).Row

    If Range("B" & i) = "" And Range("I" & i) <> "" Then
        Rows(i).Select
        Selection.Delete Shift:=xlUp
        i = i - 1
    End If
Next

lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

'-----------------總表 IPQC判定_08~10-----------------
Columns("Q:Q").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("Q1") = "IPQC判定_08~10時段"
Range("Q2").Select
ActiveCell.Formula = "=IF(C2="""","""",IF(C2=""日"",""08~10"",""20~22""))"
Selection.AutoFill Destination:=Range("Q2:Q" & lrow)
'-----------------總表 IPQC判定_08~10-----------------



'-----------------總表 IPQC判定_10~12-----------------
Columns("T:T").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("T1") = "IPQC判定_10~12時段"
Range("T2").Select
ActiveCell.Formula = "=IF(C2="""","""",IF(C2=""日"",""10~12"", ""22~24""))"
Selection.AutoFill Destination:=Range("T2:T" & lrow)
'-----------------總表 IPQC判定_10~12-----------------


'-----------------總表 IPQC判定_12~14-----------------
Columns("X:X").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("X1") = "IPQC判定_12~14時段"
Range("X2").Select
ActiveCell.Formula = "=IF(C2="""","""",IF(C2=""日"",""12~14"", ""24~02""))"
Selection.AutoFill Destination:=Range("X2:X" & lrow)
'-----------------總表 IPQC判定_12~14-----------------


'-----------------總表 IPQC判定_14~16-----------------
Columns("AD:AD").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AD1") = "IPQC判定_14~16時段"
Range("AD2").Select
ActiveCell.Formula = "=IF(C2="""","""",IF(C2=""日"", ""14~16"", ""02~04""))"
Selection.AutoFill Destination:=Range("AD2:AD" & lrow)
'-----------------總表 IPQC判定_14~16-----------------



'-----------------總表 IPQC判定_16~18-----------------
Columns("AN:AN").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AN1") = "IPQC判定_16~18時段"
Range("AN2").Select
ActiveCell.Formula = "=IF(C2="""","""",IF(C2=""日"", ""16~18"", ""04~06""))"
Selection.AutoFill Destination:=Range("AN2:AN" & lrow)
'-----------------總表 IPQC判定_16~18-----------------


'-----------------總表 IPQC判定_18~20-----------------
Columns("AP:AP").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AP1") = "IPQC判定_18~20時段"
Range("AP2").Select
ActiveCell.Formula = "=IF(C2="""","""",IF(C2=""日"", ""18~20"", ""06~08""))"
Selection.AutoFill Destination:=Range("AP2:AP" & lrow)
'-----------------總表 IPQC判定_18~20-----------------


'-----------------總表 巡檢時段-----------------
Range("AW1") = "巡檢時段"
Range("AW2").Select
ActiveCell.Formula = "=IF(P2="""",IF(S2="""",IF(W2="""",IF(AC2="""",IF(AM2="""",IF(AO2="""","""",AP2),IF(AO2="""",AN2,AN2&"";""&AP2)),IF(AM2="""",IF(AO2="""",AD2,AD2&"";""&AP2),IF(AO2="""",AD2&"";""&AN2,AD2&"";""&AN2&"";""&AP2))),IF(AC2="""",IF(AM2="""",IF(AO2="""",X2,X2&"";""&AP2),IF(AO2="""",X2&"";""&AN2,X2&"";""&AN2&"";""&AP2)),IF(AM2="""",IF(AO2="""",X2&"";""&AD2,X2&"";""&AD2&"";""&AP2),IF(AO2="""",X2&"";""&AD2&"";""&AN2,X2&"";""&AD2&"";""&AN2&"";""&AP2)))),IF(W2="""",IF(AC2="""",IF(AM2="""",IF(AO2="""",T2,T2&"";""&AP2),IF(AO2="""",T2&"";""&AN2,T2&"";""&AN2&"";""&AP2))," & _
                     "IF(AM2="""",IF(AO2="""",T2&"";""&AD2,T2&"";""&AD2&"";""&AP2),IF(AO2="""",T2&"";""&AD2&"";""&AN2,T2&"";""&AD2&"";""&AN2&"";""&AP2))),IF(AC2="""",IF(AM2="""",IF(AO2="""",T2&"";""&X2,T2&"";""&X2&"";""&AP2),IF(AO2="""",T2&"";""&X2&"";""&AN2,T2&"";""&X2&"";""&AN2&"";""&AP2)),IF(AM2="""",IF(AO2="""",T2&"";""&X2&"";""&AD2,T2&"";""&X2&"";""&AD2&"";""&AP2),IF(AO2="""",T2&"";""&X2&"";""&AD2&"";""&AN2,T2&"";""&X2&"";""&AD2&"";""&AN2&"";""&AP2))))),IF(S2="""",IF(W2="""",IF(AC2="""",IF(AM2="""",IF(AO2="""",Q2,P2&"";""&AP2),IF(AO2="""",Q2&"";""&AN2,P2&"";""&AN2&"";""&AO2))," & _
                     "IF(AM2="""",IF(AO2="""",Q2&"";""&AD2,Q2&"";""&AD2&"";""&AP2),IF(AO2="""",Q2&"";""&AD2&"";""&AN2,Q2 &"";""&AD2&"";""&AN2&"";""&AP2))),IF(AC2="""",IF(AM2="""",IF(AO2="""",Q2&"";""&X2,Q2&"";""&X2&"";""&AP2),IF(AO2="""",Q2&"";""&X2&"";""&AN2,Q2&"";""&X2&"";""&AN2&"";""&AP2)),IF(AM2="""",IF(AO2="""",Q2&"";""&X2&"";""&AD2,Q2&"";""&X2&"";""&AD2&"";""&AP2),IF(AO2="""",Q2&"";""&X2&"";""&AD2&"";""&AN2,Q2&"";""&X2&"";""&AD2&"";""&AN2&"";""&AP2)))),IF(W2="""",IF(AC2="""",IF(AM2="""",IF(AO2="""",Q2&"";""&T2,Q2&"";""&T2&"";""&AO2),IF(AN2="""",P2&"";""&T2&"";""&AN2,Q2&"";""&T2&"";""&AM2&"";""&AO2))," & _
                     "IF(AM2="""",IF(AO2="""",Q2&"";""&T2&"";""&AD2,T2&"";""&AD2&"";""&AP2),IF(AO2="""",Q2&"";""&T2&"";""&AD2&"";""&AN2,Q2&"";""&T2&"";""&AD2&"";""&AN2&"";""&AO2))),IF(AC2="""",IF(AM2="""",IF(AO2="""",Q2&"";""&T2&"";""&X2,Q2&"";""&T2&"";""&X2&"";""&AP2),IF(AO2="""",Q2&"";""&T2&"";""&X2&"";""&AN2,Q2&"";""&T2&"";""&X2&"";""&AN2&"";""&AP2)),IF(AM2="""",IF(AO2="""",Q2&"";""&T2&"";""&X2&"";""&AD2,Q2&"";""&T2&"";""&X2&"";""&AD2&"";""&AP2),IF(AO2="""",Q2&"";""&T2&"";""&X2&"";""&AD2&"";""&AN2,Q2&"";""&T2&"";""&X2&"";""&AD2&"";""&AN2&"";""&AP2))))))"
Selection.AutoFill Destination:=Range("AW2:AW" & lrow)
'-----------------總表 巡檢時段-----------------


'-----------------總表 巡檢次數-----------------
Columns("AX:AX").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AX1") = "巡檢次數"
Range("AX2").Select
ActiveCell.Formula = "=COUNTA(P2,S2,W2,AC2,AM2,AO2)"
Selection.AutoFill Destination:=Range("AX2:AX" & lrow)
'-----------------總表 巡檢次數-----------------



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
        If Range("A" & k) = Range("A" & k).Offset(-1, 0) And _
            Range("G" & k) = Range("G" & k).Offset(-1, 0) And _
            Range("I" & k) = Range("I" & k).Offset(-1, 0) Then
        
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
Range("B2", ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp)).Select
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

ActiveSheet.Range("E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("B" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------日期-----------------


'-----------------客戶-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("F2", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("C" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------客戶-----------------


'-----------------製令單號-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("I2", ActiveSheet.Range("I" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("D" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令單號-----------------



'-----------------班別-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
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
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型檢驗紀錄履歷").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------料號-----------------



'-----------------品名-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
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
ActiveSheet.Range("AR2", ActiveSheet.Range("AR" & ActiveSheet.Rows.Count).End(xlUp)).Select
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
