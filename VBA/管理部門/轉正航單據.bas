Sub �ॿ����()
'
' �ॿ���� ����
'
Set YEE = Sheets("�q��").Range("S1")

Set GEE = Sheets("OSP�ॿ����").Range("B1")

Set BEE = Sheets("OSP").Range("R1")

'
MsgBox "  *** �{�b�n�N�q����� - �ন[ ���� ] �q����ҳ�� ***  "
    Sheets("RD�q������X").Select
    Range("A2:AY600").Select
    Selection.ClearContents

    Range("A2").Select
    Sheets("�q���ॿ����").Select
    Range("A2:AY2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("RD�q������X").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    '---------------------�۰ʨ��N���誩��---------------------
    Sheets("���").Select
    Range("AT1") = "����Ƹ�"
    Range("AU1") = "���誩��"
    Range("AV1") = "�q��Ƹ�"
    Range("AW1") = "�q�檩��"
    Range("AX1") = "OSP�Ƹ�"
    Range("AY1") = "OSP����"
    Range("AZ1") = "�q��̥��誩�����D"
    Range("BA1") = "OSP�̥��誩�����D"
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    Range("AT2").Select
    ActiveCell.Formula = "=I2"
    Selection.AutoFill Destination:=Range("AT2:AT" & lrow)
    
    
    Range("AU2").Select
    ActiveCell.Formula = "=J2"
    Selection.AutoFill Destination:=Range("AU2:AU" & lrow)
    
    Range("AV2").Select
    ActiveCell.Formula = "=IF(N2="""","""",LEFT(N2,FIND(""#"",N2,1)-1))"
    Selection.AutoFill Destination:=Range("AV2:AV" & lrow)
    
    Range("AW2").Select
    ActiveCell.Formula = "=IF(N2="""","""",MID(N2,FIND(""#"",N2,1)+1,5))"
    Selection.AutoFill Destination:=Range("AW2:AW" & lrow)
    
    Range("AX2").Select
    ActiveCell.Formula = "=IF(OSP!C5="""","""",OSP!C5)"
    Selection.AutoFill Destination:=Range("AX2:AX" & lrow)
    
    Range("AY2").Select
    ActiveCell.Formula = "=IF(OSP!C5="""","""",OSP!D5)"
    Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
    
    For j = 2 To Range("AX65536").End(xlUp).Row
        If Range("AX" & j) = "" And Range("AY" & j) = "" And Range("AX" & j).Offset(1, 0) <> "" Then
            Range("AX" & j & ":AY" & j).Select
            Selection.Delete Shift:=xlUp
            j = j - 1
        End If
    Next
    
    Range("AZ2").Select
    ActiveCell.Formula = "=SUBSTITUTE(IFERROR(INDEX(A:A,MATCH(AV2,AT:AT,0),1), IF(AV2="""","""", AV2&""#""&AW2)),""#0"",""#O"",1)"
    Selection.AutoFill Destination:=Range("AZ2:AZ" & lrow)
    
    Range("BA2").Select
    ActiveCell.Formula = "=SUBSTITUTE(IFERROR(INDEX(A:A,MATCH(AX2,AT:AT,0),1), IF(AX2="""","""", AX2&""#""&AY2)),""#0"",""#O"",1)"
    Selection.AutoFill Destination:=Range("BA2:BA" & lrow)
    
    ActiveSheet.Range("AZ2", ActiveSheet.Range("AZ" & ActiveSheet.Rows.Count).End(xlUp)).Select
    Selection.Copy
    
    Worksheets("RD�q������X").Activate
    ActiveSheet.Range("AP2").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Worksheets("���").Activate
    ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
    Selection.Copy

    Worksheets("OSP�ॿ����").Activate
    ActiveSheet.Range("AR2").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    '---------------------�۰ʨ��N���誩��---------------------

        
        
    '-------��OSP
    Sheets("OSP�ॿ����").Select
    Range("C2:BA" & BEE).Select
   ' Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("RD�q������X").Select
    Range("A" & GEE).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
   '----------�C�L����t������
   Sheets("���").Select
   If Range("U1") > 1 Or Range("AF1") > 1 Then
        
      MsgBox "  ****�Y�N�C�L����t�����Ӫ�****  "
      
    ��z�������
    '�L����t��
   End If
   '---------
   MsgBox "  @@@ �Y�N��I �ॿ�����ɮ�  @@@  "

    Sheets("RD�q������X").Select

    For i = 2 To Range("A65536").End(xlUp).Row
        If Range("A" & i) = Range("A" & i).Offset(-1, 0) And Range("P" & i) = "OSP" Then

            If Left(Range("AP" & i), 1) Like "[a-z, A-Z]" Then

            Else
                Rows(i).Select
                Selection.Delete Shift:=xlUp
                i = i - 1
            End If
        End If
    Next

    Sheets("OSP�ॿ����").Range("B1").Formula = "=500-COUNTBLANK(RD�q������X!A1:A500)+1"

    Application.CutCopyMode = False
    Sheets("RD�q������X").Copy
    Sheets("RD�q������X").Select
    Sheets("RD�q������X").Name = YEE & "RD�q������X"
'    ChDir "\\YEAWEN\files-server\06_���\01_�ͺ�\��q�C���T\��q�q��P�f�ॿ��\��q�q��_�ॿ����"
'    ActiveWorkbook.SaveAs Filename:= _
'        "\\YEAWEN\files-server\06_���\01_�ͺ�\��q�C���T\��q�q��P�f�ॿ��\��q�q��_�ॿ����\" & YEE & "�ॿ��q����.xlsx" _
'        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
'    ActiveWindow.Close
'
'    Sheets("DATA").Select
'    Range("H1").Select
'
'    �ƻs�q���MARS��
    
End Sub




