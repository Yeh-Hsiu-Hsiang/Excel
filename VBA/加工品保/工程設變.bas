
Sub �u�{�]��()

    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\�]�p�ɮ׸��\�~�O��\�[�u�~�O\test�[�u�ե�_QC���������_iPad.xlsx"  '�}���ɮ�

    Dim ws As Worksheet
    Dim my_ws1, my_ws2, my_ws3, my_ws4 As String
    
    my_ws1 = "�[�uQC���綵�ت�"
    my_ws2 = "�Ͳ����`���p���R�l�ܬ���"
    my_ws3 = "�u�{�]��"
    my_ws4 = "���u�W�U"
    
    '�u�{�]��
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(my_ws3) Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
        
            Workbooks("test�[�u�ե�_QC���������_iPad.xlsx").Worksheets("�u�{�]��").Activate   '���w��e����ï�B�u�@��
            
            'Range("C1").SpecialCells(xlCellTypeLastCell)   ��̫�@�榳��ƪ���m
            Worksheets("�u�{�]��").Range("C1", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("�u�{�]��").Activate   '���w��e����ï�B�u�@��
            ActiveSheet.Range("C1", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select   '������
            Selection.Copy  '�ƻs
            
            Workbooks("test�[�u�ե�_QC���������_iPad.xlsx").Worksheets("�u�{�]��").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("C1", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            ' �Ƨ� C ��� O �檺���
            ' Key1:=Range("D1")     �̾� D ��Ƨ�
            ' Order1:=xlDescending  �����Ƨ�
            ' Header:=xlYes         �����D�C
            Columns("C:O").Sort Key1:=Range("D1"), Order1:=xlDescending, Header:=xlYes  '�̷Ӥ���Ƨ�
            
            '---------�B�zLM��
            Range("L2").Select
            Selection.Formula = "=IF($K2="""","""", CONCATENATE($K2,COUNTIF($K$1:$K2,$K2)))"  '�]�w L2�x�s�椽��
            Range("L2").Select  '���L2
            Selection.Copy  '�ƻs L2����
            
            Dim x As Integer
            x = Range("K65536").End(xlUp).Row   '�ھ�K��̫�@����ƨӧ��Ʀ@�X�C
            
            Range("L2", "L" & x).Select
            Selection.PasteSpecial  '�K�W����
            
            Range("L2", "L" & x).Select
            Selection.Copy
            Selection.PasteSpecial xlPasteValues '�u�K�W��
            
            Range("M2").Select
            Selection.Formula = "=CONCATENATE(TEXT($D2,""YYYYMMDD""),""�A"",$E2,""�A"",$O2)"   '�]�w M2�x�s���
            Range("M2").Select
            Selection.Copy
            
            Range("M2", "M" & x).Select
            Selection.PasteSpecial
            
            Range("M2", "M" & x).Select
            Selection.Copy
            Selection.PasteSpecial xlPasteValues '�u�K�W��
            '---------�B�zLM��
            
            Range("C2").Select
            
            ActiveWorkbook.Close True   '�����æs��
        End If
    Next
    
End Sub


