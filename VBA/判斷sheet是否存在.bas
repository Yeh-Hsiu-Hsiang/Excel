Attribute VB_Name = "Module1"
Sub �P�_sheet�O�_�s�b()
Attribute �P�_sheet�O�_�s�b.VB_ProcData.VB_Invoke_Func = " \n14"


    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\�]�p�ɮ׸��\�~�O��\�[�u�~�O\test.xlsx"  '�}���ɮ�


    Dim ws As Worksheet
    Dim my_ws As String
    
    my_ws = "�[�uQC���綵�ت�"
    
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(my_ws) Then   '�P�_�O�_�w�s�b�u�@��A�w�s�b���R���ª��A�ƻs
            Application.DisplayAlerts = False   '�����R���q��
            Sheets("�[�uQC���綵�ت�").Select
            ActiveWindow.SelectedSheets.Delete  '�R��Sheets
            Application.DisplayAlerts = True    '�}�ҧR���q��
            
            'Debug.Print ("already exist")
            
            Workbooks("����ï1").Activate   '���w��e����ï
            Sheets("�[�uQC���綵�ت�").Copy Before:=Workbooks("test.xlsx").Sheets(1)    '�ƻs�u�@��
            ActiveWorkbook.Close True   '�����æs��
        
        Else    '�Y���s�b�����s�W
            Workbooks("����ï1").Activate
            Sheets("�[�uQC���綵�ت�").Copy Before:=Workbooks("test.xlsx").Sheets(1)
            ActiveWorkbook.Close True
        End If
    Next
        
End Sub
