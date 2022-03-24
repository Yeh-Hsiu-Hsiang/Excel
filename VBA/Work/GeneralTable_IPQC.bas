Attribute VB_Name = "GeneralTable_IPQC"
Sub GeneralTable_IPQC()
Attribute GeneralTable_IPQC.VB_ProcData.VB_Invoke_Func = " \n14"

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

'-----------------匯出資料總表整理-----------------
Range("A:G, I:K, V:Z, AB:AF, AR:AV, BI:BM, BY:CC, CZ:DD, EP:EP, FC:FC, IM:IO").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False

Worksheets(1).Activate


'----------IPQC判定----------
Union(Range("AQ:AQ"), Range("BX:BX"), Range("CY:CY"), Range("DZ:DZ"), _
Range("FA:FA"), Range("JK:JK"), Range("KG:KG"), Range("LC:LC"), Range("LY:LY"), _
Range("MZ:MZ"), Range("RG:RG"), Range("SC:SC"), Range("SY:SY"), Range("TU:TU"), _
Range("UV:UV"), Range("ZC:ZC"), Range("ZY:ZY"), Range("AAU:AAU"), Range("ABQ:ABQ"), _
Range("ACR:ACR"), Range("AGY:AGY"), Range("AHU:AHU"), Range("AIQ:AIQ"), Range("AJM:AJM"), _
Range("AKN:AKN")).Select

Selection.Copy
Worksheets(2).Activate

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
'----------IPQC判定----------


'----------異常原因----------
Worksheets(1).Activate
Union(Range("FP:FP"), Range("NO:NO"), Range("VK:VK"), Range("ADG:ADG"), Range("ALC:ALC")).Select
Selection.Copy
Worksheets(2).Activate

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

Worksheets(2).Activate

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


Columns("A:B").Select
Selection.NumberFormatLocal = "yyyy/mm/dd"
Application.WindowState = xlNormal

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row



'-----------------匯出資料總表整理-----------------

'-----------------總表 IPQC-----------------
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("C2").Select
ActiveCell.FormulaR1C1 = "IPQC"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & lrow)
'-----------------總表 IPQC-----------------


'-----------------總表 檢驗員-----------------
Columns("AS:AS").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
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
Range("CL2").Formula = "= P2 & "" "" & AA2 & ""  "" & AL2"
Range("CL2").Select
Selection.AutoFill Destination:=Range("CL2:CL" & lrow)


'-----------------總表 (工站)作業員-----------------



'-----------------總表 IPQC抽驗數-----------------
Columns("AY:AY").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AY1") = "IPQC抽驗數"
Range("AY2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AND(AZ2>=2, AZ2<=544), 32, IF(AND(AZ2>=545, AZ2<=960), 40,  IF(AND(AZ2>=961, AZ2<=1632), 48,  IF(AND(AZ2>=1633, AZ2<=3072), 64,  IF(AZ2>=3073, 80, 1)))))"

Range("AY2").Select
Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
'-----------------總表 IPQC抽驗數-----------------



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
Columns("CO:CO").Select
Range("CO1") = "抽驗不良率"
Range("CO2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IFERROR(AW2/AZ2, 0)"

Range("CO2").Select
Selection.AutoFill Destination:=Range("CO2:CO" & lrow)
'-----------------總表 抽驗不良率-----------------



'-----------------總表 批不良率-----------------
Columns("CP:CP").Select
Range("CP1") = "批不良率"
Range("CP2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IFERROR(AW2/BA2, 0)"

Range("CP2").Select
Selection.AutoFill Destination:=Range("CP2:CP" & lrow)
'-----------------總表 批不良率-----------------



'-----------------總表 不良內容1-----------------
Columns("CD:CD").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CD1") = "不良內容1"
Range("CD2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CE2="""","""",CE2)"

Range("CD2").Select
Selection.AutoFill Destination:=Range("CD2:CD" & lrow)
'-----------------總表 不良內容1-----------------




'-----------------總表 不良內容2-----------------
Columns("CF:CF").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CF1") = "不良內容2"
Range("CF2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CG2="""","""",CG2)"

Range("CF2").Select
Selection.AutoFill Destination:=Range("CF2:CF" & lrow)
'-----------------總表 不良內容2-----------------



'-----------------總表 不良內容3-----------------
Columns("CH:CH").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CH1") = "不良內容3"
Range("CH2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CI2="""","""",CI2)"

Range("CH2").Select
Selection.AutoFill Destination:=Range("CH2:CH" & lrow)
'-----------------總表 不良內容3-----------------



'-----------------總表 不良內容4-----------------
Columns("CJ:CJ").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CJ1") = "不良內容4"
Range("CJ2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CK2="""","""",CK2)"

Range("CJ2").Select
Selection.AutoFill Destination:=Range("CJ2:CJ" & lrow)
'-----------------總表 不良內容4-----------------




'-----------------總表 不良內容5-----------------
Columns("CL:CL").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CL1") = "不良內容5"
Range("CL2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CM2="""","""",CM2)"

Range("CL2").Select
Selection.AutoFill Destination:=Range("CL2:CL" & lrow)
'-----------------總表 不良內容5-----------------



'-----------------總表 備註-----------------
Columns("CN:CN").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CN1") = "備註1"
Range("CN2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CO2="""","""",IF(CP2="""",CO2,IF(CQ2="""", CO2 & ""。  ""& CP2,IF(CR2="""",CO2 & ""。  "" & CP2 & ""。  "" & CQ2,IF(CS2="""",CO2 & ""。  "" & CP2 & ""。  "" & CQ2 & ""。  "" & CR2, CO2 & ""。  "" & CP2 & ""。  "" & CQ2 & ""。  "" & CR2 & ""。  "" & CS2)))))"

Range("CN2").Select
Selection.AutoFill Destination:=Range("CN2:CN" & lrow)
'-----------------總表 備註-----------------



'-----------------總表 判定-----------------
Columns("BE:BE").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BE1") = "判定"
Range("BE2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AND(BF2<>""NG"", BG2<>""NG"", BH2<>""NG"", BI2<>""NG"", BJ2<>""NG"", BK2<>""NG"", BL2<>""NG"", BM2<>""NG"", BN2<>""NG"", BO2<>""NG"", BP2<>""NG"", BQ2<>""NG"", BR2<>""NG"", BS2<>""NG"", BT2<>""NG"", BU2<>""NG"", BV2<>""NG"", BW2<>""NG"", BX2<>""NG"", BY2<>""NG"", BZ2<>""NG"", CA2<>""NG"", CB2<>""NG"", CC2<>""NG"", CD2<>""NG""),""OK"", ""NG"")"

Range("BE2").Select
Selection.AutoFill Destination:=Range("BE2:BE" & lrow)
'-----------------總表 判定-----------------



'-----------------總表 NG數-----------------
Columns("BE:BE").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BE1") = "NG數"
Range("BE2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=COUNTIF(BG2,""NG"") + COUNTIF(BH2,""NG"") + COUNTIF(BI2,""NG"") +COUNTIF(BJ2,""NG"") + COUNTIF(BK2,""NG"") + COUNTIF(BL2,""NG"") + COUNTIF(BM2,""NG"") + COUNTIF(BN2,""NG"") + COUNTIF(BO2,""NG"") + COUNTIF(BP2,""NG"") + COUNTIF(BQ2,""NG"") + COUNTIF(BR2,""JG"") + COUNTIF(BS2,""NG"") + COUNTIF(BT2,""NG"") + COUNTIF(BU2,""NG"") + COUNTIF(BV2,""NG"") + COUNTIF(BW2,""NG"") + COUNTIF(BX2,""NG"") + COUNTIF(BY2,""NG"") + COUNTIF(BZ2,""NG"") + COUNTIF(CA2,""NG"") + COUNTIF(CB2,""NG"") + COUNTIF(CC2,""NG"") + COUNTIF(CD2,""NG"") + COUNTIF(CE2,""NG"")"

Range("BE2").Select
Selection.AutoFill Destination:=Range("BE2:BE" & lrow)
'-----------------總表 NG數-----------------


'-----------------總表 NG數資料-----------------
For k = 2 To 5000
    If Range("BF" & k) = "NG" Then

        If Range("A" & k) = Range("A" & k).Offset(-1, 0) And Range("G" & k) = Range("G" & k).Offset(-1, 0) Then
            k = k + 1
        Else
            For m = 1 To Range("BE" & k)

                Range("A" & k & ":CX" & k).Select
                Selection.Copy

                Range("A" & k & ":CX" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next m

            Range("BF" & k) = "OK"
            Range("AX" & k & ":AY" & k) = 0

        End If
    End If
Next
'-----------------總表 NG數資料-----------------



'複製資料匯出總表 準備貼到品保 IPQC 總表
Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy
'-----------------匯出資料總表整理-----------------


'-----------------D 欄 IPQC-----------------
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
'-----------------D 欄 IPQC-----------------



'-----------------檢驗日期-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("A2", ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("E" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗日期-----------------



'-----------------檢驗員-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BB2", ActiveSheet.Range("BB" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員-----------------



'-----------------工單數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("G" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------工單數-----------------



'-----------------製令單號-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令單號-----------------



'-----------------製令日期-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("B2", ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("I" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令日期-----------------



'-----------------客戶-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("J" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------客戶-----------------



'-----------------機種-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("K" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------機種-----------------



'-----------------品名-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("F2", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("L" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------品名-----------------



'-----------------SOP-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("I2", ActiveSheet.Range("I" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("M" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------SOP-----------------



'-----------------SIP-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------SIP-----------------



'-----------------樣品-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("M2", ActiveSheet.Range("M" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------樣品-----------------



'-----------------工站數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("O2", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------工站數-----------------



'-----------------(工站)作業員-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("P2", ActiveSheet.Range("P" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------(工站)作業員-----------------



'-----------------製造數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製造數-----------------



'-----------------抽驗數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AZ2", ActiveSheet.Range("AZ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------抽驗數-----------------



'-----------------不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AW2", ActiveSheet.Range("AW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良數-----------------



'-----------------抽驗不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CW2", ActiveSheet.Range("CW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------抽驗不良率-----------------



'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CX2", ActiveSheet.Range("CX" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("W" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------批不良率-----------------



'-----------------不良內容1-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CF2", ActiveSheet.Range("CF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容1-----------------



'-----------------不良內容2-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CH2", ActiveSheet.Range("CH" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容2-----------------



'-----------------不良內容3-----------------
Workbooks(ActWb).Worksheets(2).Activate


ActiveSheet.Range("CJ2", ActiveSheet.Range("CJ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("Z" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容3-----------------



'-----------------不良內容4-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CL2", ActiveSheet.Range("CL" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("AA" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容4-----------------



'-----------------不良內容5-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CN2", ActiveSheet.Range("CN" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("AB" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良內容5-----------------



'-----------------備註1-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CP2", ActiveSheet.Range("CP" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------備註1-----------------



'-----------------判定-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BF2", ActiveSheet.Range("BF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate


ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------




End Sub
