Attribute VB_Name = "GetSheetsFromOther"
'�q���|�M�椤�}�Ҩýƻs Excel �u�@���@���ɮ�
Sub GetSheetsFromOther()

    Dim Path, ActWb As String, i As Integer
    
    ActWb = ActiveWorkbook.Name         '�����{������ï�W��
    
    For i = 1 To Workbooks(ActWb).Worksheets("test").Range("A65536").End(xlUp).Row      '�v��Ū�����|
        
        Workbooks(ActWb).Worksheets("test").Activate    '���w�ثe�u�@��
    
        Path = Range("A" & i)           '���|
        
        fileName = Dir(Path & "*.xls")      '�ɦW
        
        Do While fileName <> ""
            Workbooks.Open fileName:=Path & fileName, ReadOnly:=True        '�}�Ұ�Ū�ɮ�
            OpenWb = ActiveWorkbook.Name
            
            For x = 1 To Workbooks(OpenWb).Sheets.Count
                If Workbooks(OpenWb).Sheets.Count > 1 Then
                    ActiveWorkbook.Sheets(x).Copy _
                    After:=Workbooks(ActWb).Sheets(1)
                    ActiveSheet.Name = Range("K4") & "#" & Range("O5") & "-" & x  '�u�@��W��
                Else
                    ActiveWorkbook.Sheets(x).Copy _
                    After:=Workbooks(ActWb).Sheets(1)
                    ActiveSheet.Name = Range("K4") & "#" & Range("O5")  '�u�@��W��
                End If
            Next

            Workbooks(fileName).Close       '������Ū�ɮ�
            fileName = Dir()
          Loop
    Next i
    
End Sub


Sub Del_list()
       SendKeys "^g^a{DEL}"
End Sub

