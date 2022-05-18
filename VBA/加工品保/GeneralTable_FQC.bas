
Sub GeneralTable_FQC()

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

'-----------------匯出資料總表整理-----------------
Range("A:G, I:K, V:Z, AB:AF, AR:AV, BI:BM, BY:CC, CZ:DD, EP:EP, FC:FC, IM:IO").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(2)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False

Worksheets(1).Activate


'----------異常原因----------
Worksheets(1).Activate
Union(Range("FP:FP"), Range("NO:NO"), Range("VK:VK"), Range("ADG:ADG"), Range("ALC:ALC")).Select
Selection.Copy
Worksheets(3).Activate

i = 1
Do While True
    If ActiveSheet.Cells(1, i).Value = "" Then
        ActiveSheet.Cells(1, i).Select
        Exit Do
    End If
    i = i + 1
Loop

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
'----------異常原因----------



'----------備註----------
Worksheets(1).Activate
Range("IL:IL, QH:QH, YD:YD, AFZ:AFZ, ANV:ANV").Select

Selection.Copy

Worksheets(3).Activate

i = 1
Do While True
    If ActiveSheet.Cells(1, i).Value = "" Then
        ActiveSheet.Cells(1, i).Select
        Exit Do
    End If
    i = i + 1
Loop

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
'----------備註----------



'----------FQC----------
Worksheets(1).Activate
Union(Range("GA:GA"), Range("HW:HW"), Range("PV:PV"), Range("XR:XR"), _
Range("AFN:AFN"), Range("ANJ:ANJ")).Select

Selection.Copy
Worksheets(3).Activate

i = 1
Do While True
    If ActiveSheet.Cells(1, i).Value = "" Then
        ActiveSheet.Cells(1, i).Select
        Exit Do
    End If
    i = i + 1
Loop

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False

Columns("A:B").Select
Selection.NumberFormatLocal = "yyyy/mm/dd"
Application.WindowState = xlNormal

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row



'-----------------總表 FQC-----------------
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("C2").Select
ActiveCell.FormulaR1C1 = "FQC"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & lrow)
'-----------------總表 FQC-----------------

'-----------------總表 檢驗員-----------------
Columns("AS:AS").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AS1") = "檢驗員"
Range("AS2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(AT2=AU2, AT2, AT2 & "" "" & AU2)"
Range("AS2").Select
Selection.AutoFill Destination:=Range("AS2:AS" & lrow)
'-----------------總表 檢驗員-----------------

'-----------------總表 SOP-----------------
Columns("I:I").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("I1") = "SOP"
Range("I2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IFERROR(IF(FIND(""█"", J2)=1, ""V"", ""X""),""X"")"
Range("I2").Select
Selection.AutoFill Destination:=Range("I2:I" & lrow)
'-----------------總表 SOP-----------------


'-----------------總表 SIP-----------------
Columns("K:K").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("K1") = "SIP"
Range("K2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IFERROR(IF(FIND(""█"", L2)=1, ""V"", ""X""),""X"")"
Range("K2").Select
Selection.AutoFill Destination:=Range("K2:K" & lrow)
'-----------------總表 SIP-----------------


'-----------------總表 樣品-----------------
Columns("M:M").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("M1") = "樣品"
Range("M2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IFERROR(IF(FIND(""█"", N2)=1, ""V"", ""X""),""X"")"
Range("M2").Select
Selection.AutoFill Destination:=Range("M2:M" & lrow)
'-----------------總表 樣品-----------------



'-----------------總表 工站數-----------------
Columns("O:O").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("O1") = "工站數"
Range("O2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=COUNTA(P2:T2, Z2:AD2, AJ2:AN2)"

Range("O2").Select
Selection.AutoFill Destination:=Range("O2:O" & lrow)
'-----------------總表 工站數-----------------



'-----------------總表 (工站)作業員-----------------
Columns("P:P").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("P1") = "(工站)作業員"
Range("P2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(Q2="""","""",IF(R2="""",CONCATENATE(""("",Q2,"")"",V2),IF(S2="""",CONCATENATE(""("",Q2,"")"",V2,""  ("",R2,"")"",W2), IF(T2="""", CONCATENATE(""("", Q2, "")"", V2, ""  ("", R2, "")"", W2, ""  ("", S2, "")"", X2), IF(U2="""", CONCATENATE(""("", Q2, "")"", V2, ""  ("", R2, "")"", W2,  ""  ("", S2, "")"", X2, ""  ("", T2, "")"", Y2), CONCATENATE(""("", Q2, "")"", V2, ""  ("", R2, "")"", W2,  ""  ("", S2, "")"", X2,  ""  ("", T2, "")"", Y2,  ""  ("", U2, "")"", Z2))))))"

Range("P2").Select
Selection.AutoFill Destination:=Range("P2:P" & lrow)


'-----
Columns("AA:AA").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AA1") = "(工站)作業員"
Range("AA2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AB2="""","""",IF(AC2="""",CONCATENATE(""("",AB2,"")"",AG2),IF(AD2="""",CONCATENATE(""("",AB2,"")"",AG2,""  ("",AC2,"")"",AH2), IF(AE2="""", CONCATENATE(""("", AB2, "")"", AG2, ""  ("", AC2, "")"", AH2, ""  ("", AD2, "")"", AI2), IF(AF2="""", CONCATENATE(""("", AB2, "")"", AG2, ""  ("", AC2, "")"", AH2,  ""  ("", AD2, "")"", AI2, ""  ("", AE2, "")"", AJ2), CONCATENATE(""("", AB2, "")"", AG2, ""  ("", AC2, "")"", AH2,  ""  ("", AD2, "")"", AI2,  ""  ("", AE2, "")"", AJ2,  ""  ("", AF2, "")"", AK2))))))"

Range("AA2").Select
Selection.AutoFill Destination:=Range("AA2:AA" & lrow)


'-----
Columns("AL:AL").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AL1") = "(工站)作業員"
Range("AL2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AM2="""","""",IF(AN2="""",CONCATENATE(""("",AM2,"")"",AR2),IF(AO2="""",CONCATENATE(""("",AM2,"")"",AR2,""  ("",AN2,"")"",AS2), IF(AP2="""", CONCATENATE(""("", AM2, "")"", AR2, ""  ("", AN2, "")"", AS2, ""  ("", AO2, "")"", AT2), IF(AQ2="""", CONCATENATE(""("", AM2, "")"", AR2, ""  ("", AN2, "")"", AS2,  ""  ("", AO2, "")"", AT2, ""  ("", AP2, "")"", AU2), CONCATENATE(""("", AM2, "")"", AR2, ""  ("", AN2, "")"", AS2,  ""  ("", AO2, "")"", AT2,  ""  ("", AP2, "")"", AU2,  ""  ("", AQ2, "")"", AV2))))))"

Range("AL2").Select
Selection.AutoFill Destination:=Range("AL2:AL" & lrow)


'-----
Range("BS2").Formula = "= P2 & "" "" & AA2 & ""  "" & AL2"
Range("BS2").Select
Selection.AutoFill Destination:=Range("BS2:BS" & lrow)


'-----------------總表 (工站)作業員-----------------

'-----------------總表 FQC抽驗數-----------------
Columns("BM:BM").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BM1") = "FQC抽驗數"
Range("BM2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AND(BN2>=2, BN2<=544), 32, IF(AND(BN2>=545, BN2<=960), 40,  IF(AND(BN2>=961, BN2<=1632), 48,  IF(AND(BN2>=1633, BN2<=3072), 64,  IF(BN2>=3073, 80, 1)))))"

Range("BM2").Select
Selection.AutoFill Destination:=Range("BM2:BM" & lrow)
'-----------------總表 FQC抽驗數-----------------


'-----------------總表 不良數-----------------
Columns("AW:AW").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AW1") = "不良數總計"
Range("AW2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=SUM(AX2:AY2)"

Range("AW2").Select
Selection.AutoFill Destination:=Range("AW2:AW" & lrow)
'-----------------總表 不良數-----------------


'-----------------總表 抽驗不良率-----------------
Columns("BV:BV").Select
Range("BV1") = "抽驗不良率"
Range("BV2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IFERROR(AW2/BN2, 0)"

Range("BV2").Select
Selection.AutoFill Destination:=Range("BV2:BV" & lrow)
'-----------------總表 抽驗不良率-----------------


'-----------------總表 批不良率-----------------
Columns("BW:BW").Select
Range("BW1") = "批不良率"
Range("BW2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IFERROR(AW2/BO2, 0)"

Range("BW2").Select
Selection.AutoFill Destination:=Range("BW2:BW" & lrow)
'-----------------總表 批不良率-----------------


'-----------------總表 不良內容1-----------------
Columns("BD:BD").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BD1") = "不良內容1"
Range("BD2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BE2="""","""",BE2)"

Range("BD2").Select
Selection.AutoFill Destination:=Range("BD2:BD" & lrow)
'-----------------總表 不良內容1-----------------

'-----------------總表 不良內容2-----------------
Columns("BF:BF").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BF1") = "不良內容2"
Range("BF2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BG2="""","""",BG2)"

Range("BF2").Select
Selection.AutoFill Destination:=Range("BF2:BF" & lrow)
'-----------------總表 不良內容2-----------------


'-----------------總表 不良內容3-----------------
Columns("BH:BH").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BH1") = "不良內容3"
Range("BH2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BI2="""","""",BI2)"

Range("BH2").Select
Selection.AutoFill Destination:=Range("BH2:BH" & lrow)
'-----------------總表 不良內容3-----------------


'-----------------總表 不良內容4-----------------
Columns("BJ:BJ").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BJ1") = "不良內容4"
Range("BJ2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BK2="""","""",BK2)"

Range("BJ2").Select
Selection.AutoFill Destination:=Range("BJ2:BJ" & lrow)
'-----------------總表 不良內容4-----------------




'-----------------總表 不良內容5-----------------
Columns("BL:BL").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BL1") = "不良內容5"
Range("BL2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BM2="""","""",BM2)"

Range("BL2").Select
Selection.AutoFill Destination:=Range("BL2:BL" & lrow)
'-----------------總表 不良內容5-----------------


'-----------------總表 備註-----------------
Columns("BN:BN").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BN1") = "備註1"
Range("BN2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BO2="""","""",IF(BP2="""",BO2,IF(BQ2="""", BO2 & ""。  ""& BP2,IF(BR2="""",BO2 & ""。  "" & BP2 & ""。  "" & BQ2,IF(BS2="""",BO2 & ""。  "" & BP2 & ""。  "" & BQ2 & ""。  "" & BR2, BO2 & ""。  "" & BP2 & ""。  "" & BQ2 & ""。  "" & BR2 & ""。  "" & BS2)))))"

Range("BN2").Select
Selection.AutoFill Destination:=Range("BN2:BN" & lrow)
'-----------------總表 備註-----------------



'-----------------總表 判定-----------------
Columns("BV:BV").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BV1") = "判定"
Range("BV2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AND(BW2<>""不合格"", BX2<>""不合格"",BY2<>""不合格"", BZ2<>""不合格"", CA2<>""不合格""), ""OK"", ""NG"")"

Range("BV2").Select
Selection.AutoFill Destination:=Range("BV2:BV" & lrow)
'-----------------總表 判定-----------------


'-----------------總表 NG數-----------------
Columns("BV:BV").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BV1") = "NG數"
Range("BV2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=COUNTIF(BX2, ""不合格"")+COUNTIF(BY2, ""不合格"")+COUNTIF(BZ2, ""不合格"")+COUNTIF(CA2, ""不合格"")+COUNTIF(CB2, ""不合格"")"

Range("BV2").Select
Selection.AutoFill Destination:=Range("BV2:BV" & lrow)
'-----------------總表 NG數-----------------


'-----------------總表 NG數資料-----------------
For k = 2 To 5000
    If Range("BW" & k) = "NG" Then

        If Range("A" & k) = Range("A" & k).Offset(-1, 0) And Range("G" & k) = Range("G" & k).Offset(-1, 0) Then
            k = k + 1
        Else
            For m = 1 To Range("BV" & k)

                Range("A" & k & ":CG" & k).Select
                Selection.Copy

                Range("A" & k & ":CG" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next m

            Range("BW" & k) = "OK"
            Range("AX" & k & ":AY" & k) = 0

        End If
    End If
Next
'-----------------總表 NG數資料-----------------



'複製資料匯出總表 準備貼到品保 IPQC 總表
Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy


'-----------------匯出資料總表整理-----------------


'-----------------D 欄 FQC-----------------
Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

j = 6
Do While True
    If Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Cells(j, "D").Value = "" Then
        Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Cells(j, "D").Select
        Exit Do
    End If
    j = j + 1
Loop

ActiveSheet.Range("D" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------D 欄 FQC-----------------


'-----------------檢驗日期-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("A2", ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("E" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗日期-----------------


'-----------------檢驗員-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員-----------------


'-----------------工單數-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("G" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------工單數-----------------



'-----------------製令單號-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令單號-----------------



'-----------------製令日期-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("B2", ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("I" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令日期-----------------



'-----------------客戶-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("J" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------客戶-----------------



'-----------------機種-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("K" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------機種-----------------



'-----------------品名-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("F2", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("L" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------品名-----------------



'-----------------SOP-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("I2", ActiveSheet.Range("I" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("M" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------SOP-----------------



'-----------------SIP-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------SIP-----------------



'-----------------樣品-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("M2", ActiveSheet.Range("M" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------樣品-----------------

'-----------------工站數-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("O2", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------工站數-----------------


'-----------------(工站)作業員-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("P2", ActiveSheet.Range("P" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------(工站)作業員-----------------


'-----------------製造數-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("BU2", ActiveSheet.Range("BU" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製造數-----------------


'-----------------抽驗數-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("BT2", ActiveSheet.Range("BT" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------抽驗數-----------------


'-----------------不良數-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("AW2", ActiveSheet.Range("AW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良數-----------------


'-----------------抽驗不良率-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("CD2", ActiveSheet.Range("CD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------抽驗不良率-----------------


'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("CE2", ActiveSheet.Range("CE" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("W" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------批不良率-----------------


'-----------------不良內容1-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BD2", ActiveSheet.Range("BD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容1-----------------


'-----------------不良內容2-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BF2", ActiveSheet.Range("BF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容2-----------------


'-----------------不良內容3-----------------
Workbooks(ActWb).Worksheets(3).Activate


ActiveSheet.Range("BH2", ActiveSheet.Range("BH" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("Z" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容3-----------------


'-----------------不良內容4-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BJ2", ActiveSheet.Range("BJ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("AA" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容4-----------------


'-----------------不良內容5-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BL2", ActiveSheet.Range("BL" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("AB" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容5-----------------


'-----------------備註1-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BN2", ActiveSheet.Range("BN" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------備註1-----------------

'-----------------判定-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BW2", ActiveSheet.Range("BW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate


ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------
End Sub
