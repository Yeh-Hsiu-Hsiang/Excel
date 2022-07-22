Attribute VB_Name = "GetSheetsFromOther"
'從路徑清單中開啟並複製 Excel 工作表到一個檔案
Sub GetSheetsFromOther()

    Dim Path, ActWb As String, i As Integer
    
    ActWb = ActiveWorkbook.Name         '紀錄現有活頁簿名稱
    
    For i = 1 To Workbooks(ActWb).Worksheets("test").Range("A65536").End(xlUp).Row      '逐筆讀取路徑
        
        Workbooks(ActWb).Worksheets("test").Activate    '指定目前工作表
    
        Path = Range("A" & i)           '路徑
        
        fileName = Dir(Path & "*.xls")      '檔名
        
        Do While fileName <> ""
            Workbooks.Open fileName:=Path & fileName, ReadOnly:=True        '開啟唯讀檔案
            OpenWb = ActiveWorkbook.Name
            
            For x = 1 To Workbooks(OpenWb).Sheets.Count
                If Workbooks(OpenWb).Sheets.Count > 1 Then
                    ActiveWorkbook.Sheets(x).Copy _
                    After:=Workbooks(ActWb).Sheets(1)
                    ActiveSheet.Name = Range("K4") & "#" & Range("O5") & "-" & x  '工作表名稱
                Else
                    ActiveWorkbook.Sheets(x).Copy _
                    After:=Workbooks(ActWb).Sheets(1)
                    ActiveSheet.Name = Range("K4") & "#" & Range("O5")  '工作表名稱
                End If
            Next

            Workbooks(fileName).Close       '關閉唯讀檔案
            fileName = Dir()
          Loop
    Next i
    
End Sub


Sub Del_list()
       SendKeys "^g^a{DEL}"
End Sub

