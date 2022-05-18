
Sub GeneralTable_FQC()

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name

'-----------------�ץX����`���z-----------------
Range("A:G, I:K, V:Z, AB:AF, AR:AV, BI:BM, BY:CC, CZ:DD, EP:EP, FC:FC, IM:IO").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(2)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False

Worksheets(1).Activate


'----------���`��]----------
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
'----------���`��]----------



'----------�Ƶ�----------
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
'----------�Ƶ�----------



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



'-----------------�`�� FQC-----------------
Columns("C:C").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("C2").Select
ActiveCell.FormulaR1C1 = "FQC"
Range("C2").Select
Selection.AutoFill Destination:=Range("C2:C" & lrow)
'-----------------�`�� FQC-----------------

'-----------------�`�� �����-----------------
Columns("AS:AS").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("AS1") = "�����"
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
Range("BS2").Formula = "= P2 & "" "" & AA2 & ""  "" & AL2"
Range("BS2").Select
Selection.AutoFill Destination:=Range("BS2:BS" & lrow)


'-----------------�`�� (�u��)�@�~��-----------------

'-----------------�`�� FQC�����-----------------
Columns("BM:BM").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BM1") = "FQC�����"
Range("BM2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AND(BN2>=2, BN2<=544), 32, IF(AND(BN2>=545, BN2<=960), 40,  IF(AND(BN2>=961, BN2<=1632), 48,  IF(AND(BN2>=1633, BN2<=3072), 64,  IF(BN2>=3073, 80, 1)))))"

Range("BM2").Select
Selection.AutoFill Destination:=Range("BM2:BM" & lrow)
'-----------------�`�� FQC�����-----------------


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
Columns("BV:BV").Select
Range("BV1") = "���礣�}�v"
Range("BV2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IFERROR(AW2/BN2, 0)"

Range("BV2").Select
Selection.AutoFill Destination:=Range("BV2:BV" & lrow)
'-----------------�`�� ���礣�}�v-----------------


'-----------------�`�� �夣�}�v-----------------
Columns("BW:BW").Select
Range("BW1") = "�夣�}�v"
Range("BW2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IFERROR(AW2/BO2, 0)"

Range("BW2").Select
Selection.AutoFill Destination:=Range("BW2:BW" & lrow)
'-----------------�`�� �夣�}�v-----------------


'-----------------�`�� ���}���e1-----------------
Columns("BD:BD").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BD1") = "���}���e1"
Range("BD2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BE2="""","""",BE2)"

Range("BD2").Select
Selection.AutoFill Destination:=Range("BD2:BD" & lrow)
'-----------------�`�� ���}���e1-----------------

'-----------------�`�� ���}���e2-----------------
Columns("BF:BF").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BF1") = "���}���e2"
Range("BF2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BG2="""","""",BG2)"

Range("BF2").Select
Selection.AutoFill Destination:=Range("BF2:BF" & lrow)
'-----------------�`�� ���}���e2-----------------


'-----------------�`�� ���}���e3-----------------
Columns("BH:BH").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BH1") = "���}���e3"
Range("BH2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BI2="""","""",BI2)"

Range("BH2").Select
Selection.AutoFill Destination:=Range("BH2:BH" & lrow)
'-----------------�`�� ���}���e3-----------------


'-----------------�`�� ���}���e4-----------------
Columns("BJ:BJ").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BJ1") = "���}���e4"
Range("BJ2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BK2="""","""",BK2)"

Range("BJ2").Select
Selection.AutoFill Destination:=Range("BJ2:BJ" & lrow)
'-----------------�`�� ���}���e4-----------------




'-----------------�`�� ���}���e5-----------------
Columns("BL:BL").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BL1") = "���}���e5"
Range("BL2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BM2="""","""",BM2)"

Range("BL2").Select
Selection.AutoFill Destination:=Range("BL2:BL" & lrow)
'-----------------�`�� ���}���e5-----------------


'-----------------�`�� �Ƶ�-----------------
Columns("BN:BN").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BN1") = "�Ƶ�1"
Range("BN2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(BO2="""","""",IF(BP2="""",BO2,IF(BQ2="""", BO2 & ""�C  ""& BP2,IF(BR2="""",BO2 & ""�C  "" & BP2 & ""�C  "" & BQ2,IF(BS2="""",BO2 & ""�C  "" & BP2 & ""�C  "" & BQ2 & ""�C  "" & BR2, BO2 & ""�C  "" & BP2 & ""�C  "" & BQ2 & ""�C  "" & BR2 & ""�C  "" & BS2)))))"

Range("BN2").Select
Selection.AutoFill Destination:=Range("BN2:BN" & lrow)
'-----------------�`�� �Ƶ�-----------------



'-----------------�`�� �P�w-----------------
Columns("BV:BV").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BV1") = "�P�w"
Range("BV2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=IF(AND(BW2<>""���X��"", BX2<>""���X��"",BY2<>""���X��"", BZ2<>""���X��"", CA2<>""���X��""), ""OK"", ""NG"")"

Range("BV2").Select
Selection.AutoFill Destination:=Range("BV2:BV" & lrow)
'-----------------�`�� �P�w-----------------


'-----------------�`�� NG��-----------------
Columns("BV:BV").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("BV1") = "NG��"
Range("BV2").Select
Application.CutCopyMode = False

ActiveCell.Formula = "=COUNTIF(BX2, ""���X��"")+COUNTIF(BY2, ""���X��"")+COUNTIF(BZ2, ""���X��"")+COUNTIF(CA2, ""���X��"")+COUNTIF(CB2, ""���X��"")"

Range("BV2").Select
Selection.AutoFill Destination:=Range("BV2:BV" & lrow)
'-----------------�`�� NG��-----------------


'-----------------�`�� NG�Ƹ��-----------------
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
'-----------------�`�� NG�Ƹ��-----------------



'�ƻs��ƶץX�`�� �ǳƶK��~�O IPQC �`��
Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy


'-----------------�ץX����`���z-----------------


'-----------------D �� FQC-----------------
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
'-----------------D �� FQC-----------------


'-----------------������-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("A2", ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("E" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------������-----------------


'-----------------�����-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�����-----------------


'-----------------�u���-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("G" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�u���-----------------



'-----------------�s�O�渹-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�s�O�渹-----------------



'-----------------�s�O���-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("B2", ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("I" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�s�O���-----------------



'-----------------�Ȥ�-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("J" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�Ȥ�-----------------



'-----------------����-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("K" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------����-----------------



'-----------------�~�W-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("F2", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("L" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�~�W-----------------



'-----------------SOP-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("I2", ActiveSheet.Range("I" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("M" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------SOP-----------------



'-----------------SIP-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("N" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------SIP-----------------



'-----------------�˫~-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("M2", ActiveSheet.Range("M" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("O" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�˫~-----------------

'-----------------�u����-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("O2", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("P" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�u����-----------------


'-----------------(�u��)�@�~��-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("P2", ActiveSheet.Range("P" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("Q" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------(�u��)�@�~��-----------------


'-----------------�s�y��-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("BU2", ActiveSheet.Range("BU" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�s�y��-----------------


'-----------------�����-----------------
Workbooks(ActWb).Worksheets(3).Activate
ActiveSheet.Range("BT2", ActiveSheet.Range("BT" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�����-----------------


'-----------------���}��-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("AW2", ActiveSheet.Range("AW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}��-----------------


'-----------------���礣�}�v-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("CD2", ActiveSheet.Range("CD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���礣�}�v-----------------


'-----------------�夣�}�v-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("CE2", ActiveSheet.Range("CE" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("W" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�夣�}�v-----------------


'-----------------���}���e1-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BD2", ActiveSheet.Range("BD" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("X" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e1-----------------


'-----------------���}���e2-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BF2", ActiveSheet.Range("BF" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("Y" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e2-----------------


'-----------------���}���e3-----------------
Workbooks(ActWb).Worksheets(3).Activate


ActiveSheet.Range("BH2", ActiveSheet.Range("BH" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("Z" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e3-----------------


'-----------------���}���e4-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BJ2", ActiveSheet.Range("BJ" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("AA" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e4-----------------


'-----------------���}���e5-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BL2", ActiveSheet.Range("BL" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("AB" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}���e5-----------------


'-----------------�Ƶ�1-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BN2", ActiveSheet.Range("BN" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�Ƶ�1-----------------

'-----------------�P�w-----------------
Workbooks(ActWb).Worksheets(3).Activate

ActiveSheet.Range("BW2", ActiveSheet.Range("BW" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate


ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�P�w-----------------
End Sub
