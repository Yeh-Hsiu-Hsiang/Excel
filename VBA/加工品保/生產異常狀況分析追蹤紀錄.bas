
Sub �Ͳ����`���p���R�l�ܬ���()

    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\�]�p�ɮ׸��\�~�O��\�[�u�~�O\test�[�u�ե�_QC���������_iPad.xlsx"  '�}���ɮ�

    Dim ws As Worksheet
    Dim my_ws1, my_ws2, my_ws3, my_ws4 As String
    
    my_ws1 = "�[�uQC���綵�ت�"
    my_ws2 = "�Ͳ����`���p���R�l�ܬ���"
    my_ws3 = "�u�{�]��"
    my_ws4 = "���u�W�U"
    
    '�Ͳ����`���p���R�l�ܬ���
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(my_ws2) Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b�����ƻs�K�W
        
            Workbooks("test�[�u�ե�_QC���������_iPad.xlsx").Worksheets("�Ͳ����`���p���R�l�ܬ���").Activate   '���w��e����ï�B�u�@��
            
            Worksheets("�Ͳ����`���p���R�l�ܬ���").Range("D2", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '�M���¦����
            
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("�Ͳ����`���p���R�l�ܬ���").Activate   '���w��e����ï�B�u�@��
            ActiveSheet.Range("D1", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).Select   '����n�ƻs���d��
            Selection.Copy  '�ƻs
            
            Workbooks("test�[�u�ե�_QC���������_iPad.xlsx").Worksheets("�Ͳ����`���p���R�l�ܬ���").Activate   '��ܭn�K�W������ï�B�u�@��
            ActiveSheet.Range("D2", ActiveSheet.Range("L" & ActiveSheet.Rows.Count).End(xlUp)).Select   '����n�K�W���d��
            Selection.PasteSpecial  '�K�W
            
            ' �Ƨ� D ��� L �檺���
            ' Key1:=Range("E1")     �̾� E ��Ƨ�
            ' Order1:=xlDescending  �����Ƨ�
            ' Header:=xlYes         �����D�C
            Columns("D:L").Sort Key1:=Range("E1"), Order1:=xlDescending, Header:=xlYes  '�̷Ӥ���Ƨ�
            
            
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

