
Sub ��������_�����Ƹ����綵�ت�()

Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\03_�]�p�ɮ׸��\�~�O��\�����~�O\�������������(�g�X����)_iPad.xlsx"  '�}���ɮ�

    Dim ws As Worksheet

    '�����Ƹ����綵�ت�
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("�����Ƹ����綵�ت�") Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
        
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '���w�n�W�Ǧ�iReporter�ɮת�����ï�Τu�@��
            
            Worksheets("�����Ƹ����綵�ت�").Range("B1", ActiveSheet.Range("AT" & Range("B65536").End(xlUp).Row)).ClearContents '�M���¦����
            
            
            '�ƻs����s���B�l��Ƹ��B�~�W�W��B�Ȥ�BFa���x�B���Ǹ��B�Ҹ�
            Workbooks("20210330.xlsm").Worksheets("�����Ƹ����綵�ت�").Activate   '���w�쥻��Ƭ���ï�B�u�@��
            ActiveSheet.Range("B3", ActiveSheet.Range("H" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("B1", ActiveSheet.Range("H" & Range("B65536").End(xlUp).Row)).Select    '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            '�ƻs�޼ơB�g��
            Workbooks("20210330.xlsm").Worksheets("�����Ƹ����綵�ت�").Activate
            ActiveSheet.Range("L3", ActiveSheet.Range("M" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("I1", ActiveSheet.Range("J" & Range("B65536").End(xlUp).Row)).Select    '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            '�ƻs���
            Workbooks("20210330.xlsm").Worksheets("�����Ƹ����綵�ت�").Activate
            ActiveSheet.Range("I3", ActiveSheet.Range("I" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("K1", ActiveSheet.Range("K" & Range("B65536").End(xlUp).Row)).Select    '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            
            '�ƻs SOP�BSIP�B�зǼ�
            Workbooks("20210330.xlsm").Worksheets("�����Ƹ����綵�ت�").Activate
            ActiveSheet.Range("X3", ActiveSheet.Range("Z" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("L1", ActiveSheet.Range("N" & Range("B65536").End(xlUp).Row)).Select    '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            
            '�ƻs ���綵��
            Workbooks("20210330.xlsm").Worksheets("�����Ƹ����綵�ت�").Activate
            ActiveSheet.Range("AO3", ActiveSheet.Range("BT" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�����Ƹ����綵�ت�").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("O1", ActiveSheet.Range("AT" & Range("B65536").End(xlUp).Row)).Select    '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            
            
            Range("B1").Select
            Application.CutCopyMode = False
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
End Sub

