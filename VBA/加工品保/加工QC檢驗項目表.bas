
Sub �[�uQC���綵�ت�()

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
        If LCase(ws.Name) = LCase(my_ws1) Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
            
            Workbooks("test�[�u�ե�_QC���������_iPad.xlsx").Worksheets("�[�uQC���綵�ت�").Activate   '���w��e����ï�B�u�@��
            
            'Range("A1").SpecialCells(xlCellTypeLastCell)    �̫�@�榳��ƪ���m
            Worksheets("�[�uQC���綵�ت�").Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).ClearContents  '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("�[�uQC���綵�ت�").Activate   '���w��e����ï�B�u�@��
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select  '����n�ƻs���d��
            Selection.Copy  '�ƻs
            
            Workbooks("test�[�u�ե�_QC���������_iPad.xlsx").Worksheets("�[�uQC���綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
    '--------�K�� IPQC FQC ���������-------
    
    
    '--------�K�쭺�����������-------
    For Each ws_1 In Worksheets
        If LCase(ws_1.Name) = LCase(my_ws1) Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
            
            Workbooks("test�������������(�ե�).xlsx").Worksheets("�[�uQC���綵�ت�").Activate   '���w��e����ï�B�u�@��
            
            'Range("A1").SpecialCells(xlCellTypeLastCell)    �̫�@�榳��ƪ���m
            Worksheets("�[�uQC���綵�ت�").Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).ClearContents  '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("�[�uQC���綵�ت�").Activate   '���w��e����ï�B�u�@��
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select  '����n�ƻs���d��
            Selection.Copy  '�ƻs
            
            Workbooks("test�������������(�ե�).xlsx").Worksheets("�[�uQC���綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
    '--------�K�쭺�����������-------
End Sub

