Attribute VB_Name = "����_�����Ƹ����綵�ت�"

Sub ����_�����Ƹ����綵�ت�()

Workbooks.Open fileName:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\03_�]�p�ɮ׸��\�~�O��\�����~�O\�������������(�g�X����)_iPad.xlsx"  '�}���ɮ�
Workbooks.Open fileName:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\03_�]�p�ɮ׸��\�~�O��\�����~�O\�����g�X_QC���������_iPad.xlsx"  '�}���ɮ�


    Dim ws As Worksheet

    '--------�K�� IPQC FQC ���������-------
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("�����Ƹ����綵�ت�") Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
            
            Workbooks("�����g�X_QC���������_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '���w�n�W�Ǧ�iReporter�ɮת�����ï�Τu�@��
            
            'Range("A1").SpecialCells(xlCellTypeLastCell)    �̫�@�榳��ƪ���m
            Worksheets("�����Ƹ����綵�ت�").Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).ClearContents  '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(����).xlsm").Worksheets("�����Ƹ����綵�ت�").Activate   '���w�쥻��Ƭ���ï�B�u�@��
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select  '����n�ƻs���d��
            Selection.Copy  '�ƻs
            
            Workbooks("�����g�X_QC���������_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
    '--------�K�� IPQC FQC ���������-------
    
    

    '--------�K�쭺�����������-------
    For Each ws_1 In Worksheets
        If LCase(ws_1.Name) = LCase("�����Ƹ����綵�ت�") Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W

            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '���w�n�W�Ǧ�iReporter�ɮת�����ï�Τu�@��

            'Range("A1").SpecialCells(xlCellTypeLastCell)    �̫�@�榳��ƪ���m
            Worksheets("�����Ƹ����綵�ت�").Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).ClearContents  '�M���¦����

            Workbooks("�~�OIPQC_FQC����t��(����).xlsm").Worksheets("�����Ƹ����綵�ت�").Activate   '���w�쥻��Ƭ���ï�B�u�@��
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select  '����n�ƻs���d��
            Selection.Copy  '�ƻs

            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
    '--------�K�쭺�����������-------
    
    Application.CutCopyMode = False

End Sub

