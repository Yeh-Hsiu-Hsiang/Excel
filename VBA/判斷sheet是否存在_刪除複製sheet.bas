
Sub 判斷sheet是否存在()

    Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\設計檔案資料\品保課\加工品保\test.xlsx"  '開啟檔案

    Dim ws As Worksheet
    Dim my_ws As String
    
    my_ws = "加工QC檢驗項目表"
    
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(my_ws) Then   '判斷是否已存在工作表，已存在先刪除舊的再複製
            Application.DisplayAlerts = False   '關閉刪除通知
            Sheets("加工QC檢驗項目表").Select
            ActiveWindow.SelectedSheets.Delete  '刪除Sheets
            Application.DisplayAlerts = True    '開啟刪除通知
            
            'Debug.Print ("already exist")
            
            Workbooks("活頁簿1").Activate   '指定當前活頁簿
            Sheets("加工QC檢驗項目表").Copy Before:=Workbooks("test.xlsx").Sheets(1)    '複製工作表
            ActiveWorkbook.Close True   '關閉並存檔
        
        Else    '若不存在直接新增
            Workbooks("活頁簿1").Activate
            Sheets("加工QC檢驗項目表").Copy Before:=Workbooks("test.xlsx").Sheets(1)
            ActiveWorkbook.Close True
        End If
    Next  
End Sub
