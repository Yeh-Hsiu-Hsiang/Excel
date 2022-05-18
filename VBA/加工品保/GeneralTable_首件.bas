
Sub GeneralTable_����()
Attribute GeneralTable_����.VB_ProcData.VB_Invoke_Func = " \n14"

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name


'-----------------�ץX����`���z-----------------
Range("A:A, C:F, W:W, EX:EX, LL:LL").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False


Workbooks(ActWb).Worksheets(1).Activate

'-------------��X �P�w-------------
Range("EW:EW , LK:LK").Select
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
'-------------��X �P�w-------------


'-------------���粧�`�Ƶ�-------------
Workbooks(ActWb).Worksheets(1).Activate
Range("EZ:EZ , LN:LN").Select
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
'-------------���粧�`�Ƶ�-------------


'-----------------�ץX����`���z-----------------

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

'-----------------�`�� ����-----------------
Columns("B:B").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("B2").Select
ActiveCell.FormulaR1C1 = "����"
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B" & lrow)


Columns("A:A").Select
Selection.NumberFormatLocal = "yyyy/mm/dd"
Application.WindowState = xlNormal

Columns("G:G").Select
Selection.NumberFormatLocal = "yyyy/mm/dd"
Application.WindowState = xlNormal

'-----------------�`�� ����-----------------



'-----------------�`�� �����-----------------
Columns("H:H").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("H1") = "�����"
Range("H2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(I2=J2, I2, I2 & "" "" & J2)"
Range("H2").Select
Selection.AutoFill Destination:=Range("H2:H" & lrow)
'-----------------�`�� �����-----------------


'-----------------�`�� �s�y��-----------------
Range("O1") = "�s�y��"
Range("O2") = 1
Range("O2").Select
Selection.AutoFill Destination:=Range("O2:O" & lrow)
'-----------------�`�� �s�y��-----------------


'-----------------�`�� �����-----------------
Range("P1") = "�����"
Range("P2") = 1
Range("P2").Select
ActiveCell.Formula = "=IF(AND(O2>=2, O2<=544), 32, IF(AND(O2>=545, O2<=960), 40,  IF(AND(O2>=961, O2<=1632), 48,  IF(AND(O2>=1633, O2<=3072), 64,  IF(O2>=3073, 80, 1)))))"
Selection.AutoFill Destination:=Range("P2:P" & lrow)
'-----------------�`�� �����-----------------


'-----------------�`�� ���}��-----------------
Range("Q1") = "���}��"
Range("Q2") = 0
Range("Q2").Select
Selection.AutoFill Destination:=Range("Q2:Q" & lrow)
'-----------------�`�� ���}��-----------------


'-----------------�`�� ���礣�}�v-----------------
Range("R1") = "���礣�}�v"
Range("R2").Formula = "=IFERROR(Q2/P2, 0)"
Range("R2").Select
Selection.AutoFill Destination:=Range("R2:R" & lrow)
'-----------------�`�� ���礣�}�v-----------------



'-----------------�`�� �夣�}�v-----------------
Range("S1") = "�夣�}�v"
Range("S2").Formula = "=IFERROR(Q2/O2, 0)"
Range("S2").Select
Selection.AutoFill Destination:=Range("S2:S" & lrow)
'-----------------�`�� �夣�}�v-----------------



'-----------------�`�� ��X�P�w-----------------
Columns("K:K").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("K1") = "��X�P�w"
Range("K2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(FIND(""�i�Ͳ�"",L2)>4, ""NG"", ""OK"")"
Range("K2").Select
Selection.AutoFill Destination:=Range("K2:K" & lrow)
'-----------------�`�� ��X�P�w-----------------



'-----------------�`�� ���粧�`�Ƶ�-----------------
Columns("N:N").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("N1") = "���粧�`�Ƶ�"
Range("N2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(O2="""", """", IF(P2="""", O2, O2 & ""�C  "" & P2))"
Range("N2").Select
Selection.AutoFill Destination:=Range("N2:N" & lrow)
'-----------------�`�� ���粧�`�Ƶ�-----------------



'-----------------�`�� NG��-----------------
Range("V1") = "NG��"
Range("V2").Formula = "=COUNTIF(K2, ""NG"")"
Range("V2").Select
Selection.AutoFill Destination:=Range("V2:V" & lrow)
'-----------------�`�� NG��-----------------



'-----------------�`�� NG�Ƹ��-----------------
For k = 2 To 5000
    
    If Range("K" & k) = "NG" Then
        
        If Range("A" & k) = Range("A" & k).Offset(-1, 0) And Range("C" & k) = Range("C" & k).Offset(-1, 0) Then
            k = k + 1
        Else
            For m = 1 To Range("V" & k)

                Range("A" & k & ":V" & k).Select
                Selection.Copy

                Range("A" & k & ":V" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next m

            Range("K" & k) = "OK"
            Range("S" & k) = 0
        End If
    End If
Next
'-----------------�`�� NG�Ƹ��-----------------



'-----------------�ץX����`���z-----------------


'�ƻs��ƶץX�`�� �ǳƶK��~�O IPQC �`��
Range("B2", ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy


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

ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�����-----------------



'-----------------�s�O�渹-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�s�O�渹-----------------



'-----------------�s�O���-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
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



'-----------------�s�y��-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("Q2", ActiveSheet.Range("Q" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�s�y��-----------------



'-----------------�����-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("R2", ActiveSheet.Range("R" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�����-----------------



'-----------------���}��-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("S2", ActiveSheet.Range("S" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���}��-----------------



'-----------------���礣�}�v-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("T2", ActiveSheet.Range("T" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------���礣�}�v-----------------



'-----------------�夣�}�v-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("U2", ActiveSheet.Range("U" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("W" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�夣�}�v-----------------



'-----------------�P�w-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�P�w-----------------



'-----------------�Ƶ�1-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("N2", ActiveSheet.Range("N" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------�Ƶ�1-----------------

End Sub

