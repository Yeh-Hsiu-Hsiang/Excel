
Sub ��������_�Ͳ����`���p���R�l�ܬ���()

    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\03_�]�p�ɮ׸��\�~�O��\�����~�O\�������������(�g�X����)_iPad.xlsx"  '�}���ɮ�

    Dim ws As Worksheet

    '�Ͳ����`���p���R�l�ܬ���
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("�Ͳ����`���p���R�l�ܬ���") Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
        
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�Ͳ����`���p���R�l�ܬ���").Activate   '���w��e����ï�B�u�@��
            
            Worksheets("�Ͳ����`���p���R�l�ܬ���").Range("A1", ActiveSheet.Range("AJ" & Range("A65536").End(xlUp).Row)).ClearContents '�M���¦����
            
            Workbooks("20210330.xlsm").Worksheets("�Ͳ����`���p���R�l�ܬ���").Activate   '���w��e����ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("AJ" & Range("D65536").End(xlUp).Row)).Select   '����n�ƻs���d��
            Selection.Copy  '�ƻs
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�Ͳ����`���p���R�l�ܬ���").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("AJ" & Range("D65536").End(xlUp).Row)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            ' �Ƨ� D ��� L �檺���
            ' Key1:=Range("E1")     �̾� E ��Ƨ�
            ' Order1:=xlDescending  �����Ƨ�
            ' Header:=xlYes         �����D�C
            Columns("A:AJ").Sort Key1:=Range("E1"), Order1:=xlDescending, Header:=xlYes  '�̷Ӥ���Ƨ�
            
            
            '---------�B�zAB��
            Range("A2").Select
            Selection.Formula = "=IF($D2="""","""", D2&COUNTIF($D$1:$D2,$D2))"  '�]�w A2�x�s�椽��
            Range("A2").Select  '���A2
            Selection.Copy  '�ƻs A2����
            
            Dim x As Integer
            x = Range("D65536").End(xlUp).Row   '�ھ�D��̫�@����ƨӧ��Ʀ@�X�C
            
            Range("A2", "A" & x).Select
            Selection.PasteSpecial  '�K�W����
            
            Range("A2", "A" & x).Select
            Selection.Copy
            Selection.PasteSpecial xlPasteValues '�u�K�W��
            
            Range("B2").Select
            Selection.Formula = "=CONCATENATE(TEXT($E2,""YYYYMMDD""),""�A"",$H2)"   '�]�w B2�x�s���
            Range("B2").Select
            Selection.Copy
            
            Range("B2", "B" & x).Select
            Selection.PasteSpecial
            
            Range("B2", "B" & x).Select
            Selection.Copy
            Selection.PasteSpecial xlPasteValues '�u�K�W��
            '---------�B�zAB��
            
            Range("A1").Select
            
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next


End Sub
