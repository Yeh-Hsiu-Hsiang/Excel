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
      
      
    �L����t��
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
    ChDir "\\YEAWEN\files-server\06_���\01_�ͺ�\��q�C���T\��q�q��P�f�ॿ��\��q�q��_�ॿ����"
    ActiveWorkbook.SaveAs Filename:= _
        "\\YEAWEN\files-server\06_���\01_�ͺ�\��q�C���T\��q�q��P�f�ॿ��\��q�q��_�ॿ����\" & YEE & "�ॿ��q����.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close

    Sheets("DATA").Select
    Range("H1").Select

    �ƻs�q���MARS��
    
End Sub


