
Sub ��������_�M�Ʈƺ޺ި��()

Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\03_�]�p�ɮ׸��\�~�O��\�����~�O\�������������(�g�X����)_iPad.xlsx"  '�}���ɮ�

    Dim ws As Worksheet

    '�M�Ʈƺ޺ި��
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("�M�Ʈƺ޺ި��") Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
        
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�M�Ʈƺ޺ި��").Activate   '���w�n�W�Ǧ�iReporter�ɮת�����ï�Τu�@��
            
            Worksheets("�M�Ʈƺ޺ި��").Range("A1", ActiveSheet.Range("E" & Range("A65536").End(xlUp).Row)).ClearContents '�M���¦����
            
            Workbooks("20210330.xlsm").Worksheets("�M�Ʈƺ޺ި��").Activate   '���w�쥻��Ƭ���ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("E" & Range("A65536").End(xlUp).Row)).Select   '������
            Selection.Copy  '�ƻs
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�M�Ʈƺ޺ި��").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("E" & Range("A65536").End(xlUp).Row)).Select    '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            Range("A4").Select
            Application.CutCopyMode = False
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
End Sub
