Attribute VB_Name = "�P�_sheet�O�_�s�b"
Sub �P�_sheet�O�_�s�b()

    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\�]�p�ɮ׸��\�~�O��\�[�u�~�O\test.xlsx"  '�}���ɮ�

    Dim ws As Worksheet
    Dim my_ws1, my_ws2, my_ws3, my_ws4 As String
    
    my_ws1 = "�[�uQC���綵�ت�"
    my_ws2 = "�Ͳ����`���p���R�l�ܬ���"
    my_ws3 = "�u�{�]��"
    my_ws4 = "���u�W�U"
    
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(my_ws2) Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
            
            Worksheets("�Ͳ����`���p���R�l�ܬ���").Range("D2", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("�Ͳ����`���p���R�l�ܬ���").Activate   '���w��e����ï�B�u�@��
            ActiveSheet.Range("D1", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).Select   '����n�ƻs���d��
            Selection.Copy  '�ƻs
            Workbooks("test.xlsx").Worksheets("�Ͳ����`���p���R�l�ܬ���").Activate   '��ܭn�K�W����m
            ActiveSheet.Range("D2", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            ' �Ƨ� D ��� L �檺���
            ' Key1:=Range("E1")     �̾� E ��Ƨ�
            ' Order1:=xlDescending  �����Ƨ�
            ' Header:=xlYes         �����D�C
            Columns("D:L").Sort Key1:=Range("E1"), Order1:=xlDescending, Header:=xlYes  '�̷Ӥ���Ƨ�
            
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
End Sub
