Attribute VB_Name = "���u�W�U"
Sub ���u�W�U()

    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\�]�p�ɮ׸��\�~�O��\�[�u�~�O\test�[�u�ե�_QC���������_iPad.xlsx"  '�}���ɮ�
    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\�]�p�ɮ׸��\�~�O��\�[�u�~�O\test�������������(�ե�).xlsx"  '�}���ɮ�

    Dim ws, ws_1 As Worksheet
    Dim my_ws1, my_ws2, my_ws3, my_ws4 As String
    
    my_ws1 = "�[�uQC���綵�ت�"
    my_ws2 = "�Ͳ����`���p���R�l�ܬ���"
    my_ws3 = "�u�{�]��"
    my_ws4 = "���u�W�U"
    
    '--------�K�� IPQC FQC ���������-------
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(my_ws4) Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
            
            Workbooks("test�[�u�ե�_QC���������_iPad.xlsx").Worksheets("���u�W�U").Activate   '���w��e����ï�B�u�@��
            Worksheets("���u�W�U").Range("A1", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("���u�W�U").Activate   '���w��e����ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select   '������
            Selection.Copy  '�ƻs
            
            Workbooks("test�[�u�ե�_QC���������_iPad.xlsx").Worksheets("���u�W�U").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            ' �Ƨ� A ��� F �檺���
            ' Key1:=Range("B1")     �̾� B ��Ƨ�
            ' Order1:=xlAscending  �ɾ��Ƨ�
            ' Header:=xlYes         �����D�C
            Columns("A:C").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes  '�̷ӽs���Ƨ�
            
            Range("A1").Select
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
    '--------�K�� IPQC FQC ���������-------
    
    '--------�K�쭺�����������-------
    For Each ws_1 In Worksheets
        If LCase(ws_1.Name) = LCase(my_ws4) Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
            Workbooks("test�������������(�ե�).xlsx").Worksheets("���u�W�U").Activate   '���w��e����ï�B�u�@��
            Worksheets("���u�W�U").Range("A1", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("���u�W�U").Activate   '���w��e����ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select   '������
            Selection.Copy  '�ƻs
            
            Workbooks("test�������������(�ե�).xlsx").Worksheets("���u�W�U").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            ' �Ƨ� A ��� F �檺���
            ' Key1:=Range("B1")     �̾� B ��Ƨ�
            ' Order1:=xlAscending  �ɾ��Ƨ�
            ' Header:=xlYes         �����D�C
            Columns("A:C").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes  '�̷ӽs���Ƨ�
            
            Range("A1").Select
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
    '--------�K�쭺�����������-------
End Sub



