Attribute VB_Name = "����_�u�{�]��"
Sub ����_�u�{�]��()

    Workbooks.Open fileName:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\03_�]�p�ɮ׸��\�~�O��\�����~�O\�������������(�g�X����)_iPad.xlsx"  '�}���ɮ�
    Workbooks.Open fileName:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\03_�]�p�ɮ׸��\�~�O��\�����~�O\�����g�X_QC���������_iPad.xlsx"  '�}���ɮ�
    Workbooks.Open fileName:="\\yeawen\files-server\05_�~�O\13-3�Ӳե�(�׬�)\�~�OIPQC_FQC����t��(�ե�20210305.xlsm"

    Dim ws As Worksheet

    '�u�{�]��
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("�u�{�]��") Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
        
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�u�{�]��").Activate   '���w���󪺬���ï�Τu�@��
            Worksheets("�u�{�]��").Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).ClearContents '�M���¦����
            
            Workbooks("�����g�X_QC���������_iPad.xlsx").Worksheets("�u�{�]��").Activate   '���wQC�������������ï�Τu�@��
            Worksheets("�u�{�]��").Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).ClearContents '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("�u�{�]��").Activate   '���w�쥻��Ƭ���ï�B�u�@��
            ActiveSheet.Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).Select   '������
            Selection.Copy  '�ƻs
            
            Workbooks("�~�OIPQC_FQC����t��(����).xlsm").Worksheets("�u�{�]��").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�u�{�]��").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            Workbooks("�����g�X_QC���������_iPad.xlsx").Worksheets("�u�{�]��").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            
            '-------------------------�����g�X_QC���������-------------------------
            
            ' �Ƨ� C ��� P �檺���
            ' Key1:=Range("D1")     �̾� D ��Ƨ�
            ' Order1:=xlDescending  �����Ƨ�
            ' Header:=xlYes         �����D�C
            Columns("C:P").Sort Key1:=Range("D1"), Order1:=xlDescending, Header:=xlYes  '�̷Ӥ���Ƨ�
            
'            '---------�B�zLM��
'            Range("L2").Select
'            Selection.Formula = "=IF($K2="""","""", CONCATENATE($K2,COUNTIF($K$1:$K2,$K2)))"  '�]�w L2�x�s�椽��
'            Range("L2").Select  '���L2
'            Selection.Copy  '�ƻs L2����
'
'            Dim x As Integer
'            x = Range("D65536").End(xlUp).Row   '�ھ� D ��̫�@����ƨӧ��Ʀ@�X�C
'
'            Range("L2", "L" & x).Select
'            Selection.PasteSpecial  '�K�W����
'
'            Range("L2", "L" & x).Select
'            Selection.Copy
'            Selection.PasteSpecial xlPasteValues '�u�K�W��
'
'            Range("M2").Select
'            Selection.Formula = "=CONCATENATE(TEXT($D2,""YYYYMMDD""),""�A"",$E2,""�A"",$O2)"   '�]�w M2�x�s���
'            Range("M2").Select
'            Selection.Copy
'
'            Range("M2", "M" & x).Select
'            Selection.PasteSpecial
'
'            Range("M2", "M" & x).Select
'            Selection.Copy
'            Selection.PasteSpecial xlPasteValues '�u�K�W��
            '---------�B�zLM��
            
            Range("D2").Select
            
            '-------------------------�����g�X_QC���������-------------------------
            
            
            '-------------------------�������������-------------------------
            Workbooks("�������������(�g�X����)_iPad.xlsx").Worksheets("�u�{�]��").Activate   '��ܬ���ï�B�u�@��
            
            Columns("C:P").Sort Key1:=Range("D1"), Order1:=xlDescending, Header:=xlYes  '�̷Ӥ���Ƨ�
            
'            '---------�B�zLM��
'            Range("L2").Select
'            Selection.Formula = "=IF($K2="""","""", CONCATENATE($K2,COUNTIF($K$1:$K2,$K2)))"  '�]�w L2�x�s�椽��
'            Range("L2").Select  '���L2
'            Selection.Copy  '�ƻs L2����
'
'            Dim y As Integer
'            y = Range("D65536").End(xlUp).Row   '�ھ� D ��̫�@����ƨӧ��Ʀ@�X�C
'
'            Range("L2", "L" & y).Select
'            Selection.PasteSpecial  '�K�W����
'
'            Range("L2", "L" & y).Select
'            Selection.Copy
'            Selection.PasteSpecial xlPasteValues '�u�K�W��
'
'            Range("M2").Select
'            Selection.Formula = "=CONCATENATE(TEXT($D2,""YYYYMMDD""),""�A"",$E2,""�A"",$O2)"   '�]�w M2�x�s���
'            Range("M2").Select
'            Selection.Copy
'
'            Range("M2", "M" & y).Select
'            Selection.PasteSpecial
'
'            Range("M2", "M" & y).Select
'            Selection.Copy
'            Selection.PasteSpecial xlPasteValues '�u�K�W��
'            '---------�B�zLM��
            
            Range("D2").Select
            
            '-------------------------�������������-------------------------
            
            Application.CutCopyMode = False
            Workbooks("�������������(�g�X����)_iPad.xlsx").Close True   '�����æs��
            Workbooks("�����g�X_QC���������_iPad.xlsx").Close True   '�����æs��
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Close False
        End If
    Next
    
End Sub



