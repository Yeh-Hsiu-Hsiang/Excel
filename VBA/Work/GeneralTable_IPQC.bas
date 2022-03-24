Attribute VB_Name = "GeneralTable_IPQC"
Sub GeneralTable_IPQC()
Attribute GeneralTable_IPQC.VB_ProcData.VB_Invoke_Func = " \n14"

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

'-----------------�ץX����`���z-----------------
Range("A:G, I:K, V:Z, AB:AF, AR:AV, BI:BM, BY:CC, CZ:DD, EP:EP, FC:FC, IM:IO").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False

Worksheets(1).Activate


'----------IPQC�P�w----------
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
'----------IPQC�P�w----------


'----------���`��]----------
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
'----------���`��]----------



'----------�Ƶ�----------
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
'----------�Ƶ�----------


Columns("A:B").Select
Selection.NumberFormatLocal = "yyyy/mm/dd"
Application.WindowState = xlNormal

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row



'-----------------�ץX����`���z-----------------

'-----------------�`�� IPQC-----------------
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("C2").Select
ActiveCell.FormulaR1C1 = "IPQC"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & lrow)
'-----------------�`�� IPQC-----------------


'-----------------�`�� �����-----------------
Columns("AS:AS").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AS2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(AT2=AU2, AT2, AT2 & "" "" & AU2)"
Range("AS2").Select
Selection.AutoFill Destination:=Range("AS2:AS" & lrow)
'-----------------�`�� �����-----------------


'-----------------�`�� SOP-----------------
Columns("I:I").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("I1") = "SOP"
Range("I2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IFERROR(IF(FIND(""�i"", J2)=1, ""V"", ""X""),""X"")"
Range("I2").Select
Selection.AutoFill Destination:=Range("I2:I" & lrow)
'-----------------�`�� SOP-----------------


'-----------------�`�� SIP-----------------
Columns("K:K").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("K1") = "SIP"
Range("K2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IFERROR(IF(FIND(""�i"", L2)=1, ""V"", ""X""),""X"")"
Range("K2").Select
Selection.AutoFill Destination:=Range("K2:K" & lrow)
'-----------------�`�� SIP-----------------


'-----------------�`�� �˫~-----------------
Columns("M:M").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("M1") = "�˫~"
Range("M2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IFERROR(IF(FIND(""�i"", N2)=1, ""V"", ""X""),""X"")"
Range("M2").Select
Selection.AutoFill Destination:=Range("M2:M" & lrow)
'-----------------�`�� �˫~-----------------



'-----------------�`�� �u����-----------------
Columns("O:O").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("O1") = "�u����"
Range("O2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=COUNTA(P2:T2, Z2:AD2, AJ2:AN2)"

Range("O2").Select
Selection.AutoFill Destination:=Range("O2:O" & lrow)
'-----------------�`�� �u����-----------------



'-----------------�`�� (�u��)�@�~��-----------------
Columns("P:P").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("P1") = "(�u��)�@�~��"
Range("P2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(Q2="""","""",IF(R2="""",CONCATENATE(""("",Q2,"")"",V2),IF(S2="""",CONCATENATE(""("",Q2,"")"",V2,""  ("",R2,"")"",W2), IF(T2="""", CONCATENATE(""("", Q2, "")"", V2, ""  ("", R2, "")"", W2, ""  ("", S2, "")"", X2), IF(U2="""", CONCATENATE(""("", Q2, "")"", V2, ""  ("", R2, "")"", W2,  ""  ("", S2, "")"", X2, ""  ("", T2, "")"", Y2), CONCATENATE(""("", Q2, "")"", V2, ""  ("", R2, "")"", W2,  ""  ("", S2, "")"", X2,  ""  ("", T2, "")"", Y2,  ""  ("", U2, "")"", Z2))))))"

Range("P2").Select
Selection.AutoFill Destination:=Range("P2:P" & lrow)


'-----
Columns("AA:AA").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AA1") = "(�u��)�@�~��"
Range("AA2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AB2="""","""",IF(AC2="""",CONCATENATE(""("",AB2,"")"",AG2),IF(AD2="""",CONCATENATE(""("",AB2,"")"",AG2,""  ("",AC2,"")"",AH2), IF(AE2="""", CONCATENATE(""("", AB2, "")"", AG2, ""  ("", AC2, "")"", AH2, ""  ("", AD2, "")"", AI2), IF(AF2="""", CONCATENATE(""("", AB2, "")"", AG2, ""  ("", AC2, "")"", AH2,  ""  ("", AD2, "")"", AI2, ""  ("", AE2, "")"", AJ2), CONCATENATE(""("", AB2, "")"", AG2, ""  ("", AC2, "")"", AH2,  ""  ("", AD2, "")"", AI2,  ""  ("", AE2, "")"", AJ2,  ""  ("", AF2, "")"", AK2))))))"

Range("AA2").Select
Selection.AutoFill Destination:=Range("AA2:AA" & lrow)


'-----
Columns("AL:AL").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AL1") = "(�u��)�@�~��"
Range("AL2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AM2="""","""",IF(AN2="""",CONCATENATE(""("",AM2,"")"",AR2),IF(AO2="""",CONCATENATE(""("",AM2,"")"",AR2,""  ("",AN2,"")"",AS2), IF(AP2="""", CONCATENATE(""("", AM2, "")"", AR2, ""  ("", AN2, "")"", AS2, ""  ("", AO2, "")"", AT2), IF(AQ2="""", CONCATENATE(""("", AM2, "")"", AR2, ""  ("", AN2, "")"", AS2,  ""  ("", AO2, "")"", AT2, ""  ("", AP2, "")"", AU2), CONCATENATE(""("", AM2, "")"", AR2, ""  ("", AN2, "")"", AS2,  ""  ("", AO2, "")"", AT2,  ""  ("", AP2, "")"", AU2,  ""  ("", AQ2, "")"", AV2))))))"

Range("AL2").Select
Selection.AutoFill Destination:=Range("AL2:AL" & lrow)


'-----
Range("CL2").Formula = "= P2 & "" "" & AA2 & ""  "" & AL2"
Range("CL2").Select
Selection.AutoFill Destination:=Range("CL2:CL" & lrow)


'-----------------�`�� (�u��)�@�~��-----------------



'-----------------�`�� IPQC�����-----------------
Columns("AY:AY").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AY1") = "IPQC�����"
Range("AY2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AND(AZ2>=2, AZ2<=544), 32, IF(AND(AZ2>=545, AZ2<=960), 40,  IF(AND(AZ2>=961, AZ2<=1632), 48,  IF(AND(AZ2>=1633, AZ2<=3072), 64,  IF(AZ2>=3073, 80, 1)))))"

Range("AY2").Select
Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
'-----------------�`�� IPQC�����-----------------



'-----------------�`�� ���}��-----------------
Columns("AW:AW").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AW1") = "���}���`�p"
Range("AW2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=SUM(AX2:AY2)"

Range("AW2").Select
Selection.AutoFill Destination:=Range("AW2:AW" & lrow)
'-----------------�`�� ���}��-----------------



'-----------------�`�� ���礣�}�v-----------------
Columns("CO:CO").Select
Range("CO1") = "���礣�}�v"
Range("CO2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IFERROR(AW2/AZ2, 0)"

Range("CO2").Select
Selection.AutoFill Destination:=Range("CO2:CO" & lrow)
'-----------------�`�� ���礣�}�v-----------------



'-----------------�`�� �夣�}�v-----------------
Columns("CP:CP").Select
Range("CP1") = "�夣�}�v"
Range("CP2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IFERROR(AW2/BA2, 0)"

Range("CP2").Select
Selection.AutoFill Destination:=Range("CP2:CP" & lrow)
'-----------------�`�� �夣�}�v-----------------



'-----------------�`�� ���}���e1-----------------
Columns("CD:CD").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CD1") = "���}���e1"
Range("CD2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CE2="""","""",CE2)"

Range("CD2").Select
Selection.AutoFill Destination:=Range("CD2:CD" & lrow)
'-----------------�`�� ���}���e1-----------------




'-----------------�`�� ���}���e2-----------------
Columns("CF:CF").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CF1") = "���}���e2"
Range("CF2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CG2="""","""",CG2)"

Range("CF2").Select
Selection.AutoFill Destination:=Range("CF2:CF" & lrow)
'-----------------�`�� ���}���e2-----------------



'-----------------�`�� ���}���e3-----------------
Columns("CH:CH").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CH1") = "���}���e3"
Range("CH2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CI2="""","""",CI2)"

Range("CH2").Select
Selection.AutoFill Destination:=Range("CH2:CH" & lrow)
'-----------------�`�� ���}���e3-----------------



'-----------------�`�� ���}���e4-----------------
Columns("CJ:CJ").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CJ1") = "���}���e4"
Range("CJ2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CK2="""","""",CK2)"

Range("CJ2").Select
Selection.AutoFill Destination:=Range("CJ2:CJ" & lrow)
'-----------------�`�� ���}���e4-----------------




'-----------------�`�� ���}���e5-----------------
Columns("CL:CL").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CL1") = "���}���e5"
Range("CL2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CM2="""","""",CM2)"

Range("CL2").Select
Selection.AutoFill Destination:=Range("CL2:CL" & lrow)
'-----------------�`�� ���}���e5-----------------



'-----------------�`�� �Ƶ�-----------------
Columns("CN:CN").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("CN1") = "�Ƶ�1"
Range("CN2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(CO2="""","""",IF(CP2="""",CO2,IF(CQ2="""", CO2 & ""�C  ""& CP2,IF(CR2="""",CO2 & ""�C  "" & CP2 & ""�C  "" & CQ2,IF(CS2="""",CO2 & ""�C  "" & CP2 & ""�C  "" & CQ2 & ""�C  "" & CR2, CO2 & ""�C  "" & CP2 & ""�C  "" & CQ2 & ""�C  "" & CR2 & ""�C  "" & CS2)))))"

Range("CN2").Select
Selection.AutoFill Destination:=Range("CN2:CN" & lrow)
'-----------------�`�� �Ƶ�-----------------



'-----------------�`�� �P�w-----------------
Columns("BE:BE").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BE1") = "�P�w"
Range("BE2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AND(BF2<>""NG"", BG2<>""NG"", BH2<>""NG"", BI2<>""NG"", BJ2<>""NG"", BK2<>""NG"", BL2<>""NG"", BM2<>""NG"", BN2<>""NG"", BO2<>""NG"", BP2<>""NG"", BQ2<>""NG"", BR2<>""NG"", BS2<>""NG"", BT2<>""NG"", BU2<>""NG"", BV2<>""NG"", BW2<>""NG"", BX2<>""NG"", BY2<>""NG"", BZ2<>""NG"", CA2<>""NG"", CB2<>""NG"", CC2<>""NG"", CD2<>""NG""),""OK"", ""NG"")"

Range("BE2").Select
Selection.AutoFill Destination:=Range("BE2:BE" & lrow)
'-----------------�`�� �P�w-----------------



'-----------------�`�� NG��-----------------
Columns("BE:BE").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BE1") = "NG��"
Range("BE2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=COUNTIF(BG2,""NG"") + COUNTIF(BH2,""NG"") + COUNTIF(BI2,""NG"") +COUNTIF(BJ2,""NG"") + COUNTIF(BK2,""NG"") + COUNTIF(BL2,""NG"") + COUNTIF(BM2,""NG"") + COUNTIF(BN2,""NG"") + COUNTIF(BO2,""NG"") + COUNTIF(BP2,""NG"") + COUNTIF(BQ2,""NG"") + COUNTIF(BR2,""JG"") + COUNTIF(BS2,""NG"") + COUNTIF(BT2,""NG"") + COUNTIF(BU2,""NG"") + COUNTIF(BV2,""NG"") + COUNTIF(BW2,""NG"") + COUNTIF(BX2,""NG"") + COUNTIF(BY2,""NG"") + COUNTIF(BZ2,""NG"") + COUNTIF(CA2,""NG"") + COUNTIF(CB2,""NG"") + COUNTIF(CC2,""NG"") + COUNTIF(CD2,""NG"") + COUNTIF(CE2,""NG"")"

Range("BE2").Select
Selection.AutoFill Destination:=Range("BE2:BE" & lrow)
'-----------------�`�� NG��-----------------


'-----------------�`�� NG�Ƹ��-----------------
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
'-----------------�`�� NG�Ƹ��-----------------



'�ƻs��ƶץX�`�� �ǳƶK��~�O IPQC �`��
Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy
'-----------------�ץX����`���z-----------------


'-----------------D �� IPQC-----------------
Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

j = 6
Do While True
    If Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Cells(j, "D").Value = "" Then
        Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Cells(j, "D").Select
        Exit Do
    End If
    j = j + 1
Loop

ActiveSheet.Range("D" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------D �� IPQC-----------------



'-----------------������-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("A2", ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("E" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------������-----------------



'-----------------�����-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BB2", ActiveSheet.Range("BB" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�����-----------------



'-----------------�u���-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("G" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�u���-----------------



'-----------------�s�O�渹-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�s�O�渹-----------------



'-----------------�s�O���-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("B2", ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("I" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�s�O���-----------------



'-----------------�Ȥ�-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("J" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�Ȥ�-----------------



'-----------------����-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("K" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------����-----------------



'-----------------�~�W-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("F2", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("L" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�~�W-----------------



'-----------------SOP-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("I2", ActiveSheet.Range("I" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("M" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------SOP-----------------



'-----------------SIP-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------SIP-----------------



'-----------------�˫~-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("M2", ActiveSheet.Range("M" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�˫~-----------------



'-----------------�u����-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("O2", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�u����-----------------



'-----------------(�u��)�@�~��-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("P2", ActiveSheet.Range("P" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------(�u��)�@�~��-----------------



'-----------------�s�y��-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�s�y��-----------------



'-----------------�����-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("AZ2", ActiveSheet.Range("AZ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�����-----------------



'-----------------���}��-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("AW2", ActiveSheet.Range("AW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}��-----------------



'-----------------���礣�}�v-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CW2", ActiveSheet.Range("CW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���礣�}�v-----------------



'-----------------�夣�}�v-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CX2", ActiveSheet.Range("CX" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("W" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�夣�}�v-----------------



'-----------------���}���e1-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CF2", ActiveSheet.Range("CF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e1-----------------



'-----------------���}���e2-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CH2", ActiveSheet.Range("CH" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e2-----------------



'-----------------���}���e3-----------------
Workbooks(ActWb).Worksheets(2).Activate


ActiveSheet.Range("CJ2", ActiveSheet.Range("CJ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("Z" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e3-----------------



'-----------------���}���e4-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CL2", ActiveSheet.Range("CL" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("AA" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e4-----------------



'-----------------���}���e5-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CN2", ActiveSheet.Range("CN" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("AB" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e5-----------------



'-----------------�Ƶ�1-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("CP2", ActiveSheet.Range("CP" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�Ƶ�1-----------------



'-----------------�P�w-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("BF2", ActiveSheet.Range("BF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate


ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�P�w-----------------




End Sub
