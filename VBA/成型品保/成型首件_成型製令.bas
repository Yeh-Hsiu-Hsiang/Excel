
Sub ��������_�����s�O()

Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\03_�]�p�ɮ׸��\�~�O��\�����~�O\�������������(�g�X����)_iPad.xlsx"  '�}���ɮ�

    Dim ws As Worksheet

    '�����s�O
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("�����s�O") Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
        
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����s�O").Activate   '���w�n�W�Ǧ�iReporter�ɮת�����ï�Τu�@��
            
            Worksheets("�����s�O").Range("A1", ActiveSheet.Range("S" & Range("A65536").End(xlUp).Row)).ClearContents '�M���¦����
            
            Workbooks("20210330.xlsm").Worksheets("�����s�O").Activate   '���w�쥻��Ƭ���ï�B�u�@��
            ActiveSheet.Range("A2", ActiveSheet.Range("S" & Range("A65536").End(xlUp).Row)).Select   '������
            Selection.Copy  '�ƻs
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����s�O").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("S" & Range("A65536").End(xlUp).Row)).Select    '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            Range("A1").Select
            Application.CutCopyMode = False
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
End Sub

