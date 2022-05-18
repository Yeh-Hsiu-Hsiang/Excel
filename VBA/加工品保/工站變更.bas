Attribute VB_Name = "工站變更"
Sub 工站變更()

    Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\設計檔案資料\品保課\加工品保\test加工組立_QC檢驗紀錄表_iPad.xlsx"  '開啟檔案
    Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\設計檔案資料\品保課\加工品保\test首件檢驗紀錄表(組立).xlsx"  '開啟檔案

    Dim ws, ws_1 As Worksheet
    Dim my_ws1, my_ws2, my_ws3, my_ws4 As String
    
    my_ws1 = "加工QC檢驗項目表"
    my_ws2 = "生產異常狀況分析追蹤紀錄"
    my_ws3 = "工程設變"
    my_ws4 = "員工名冊"
    
    '--------貼到 IPQC FQC 檢驗紀錄表-------
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(my_ws4) Then   '判斷是否已存在工作表，已存在直接複製貼上
        
            Workbooks("test加工組立_QC檢驗紀錄表_iPad.xlsx").Worksheets("員工名冊").Activate   '指定當前活頁簿、工作表
            Worksheets("員工名冊").Range("I1", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有工站資料
            
            Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("員工名冊").Activate   '指定當前活頁簿、工作表
            ActiveSheet.Range("I1", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).Select   '選取資料
            Selection.Copy  '複製
            
            Workbooks("test加工組立_QC檢驗紀錄表_iPad.xlsx").Worksheets("員工名冊").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("I1", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).Select   '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            ' 排序 I 欄至 J 欄的資料
            ' Key1:=Range("I1")     依據 I 欄排序
            ' Order1:=xlAscending  升冪排序
            ' Header:=xlYes         有標題列
            Columns("I:J").Sort Key1:=Range("I1"), Order1:=xlAscending, Header:=xlYes  '依照編號排序
            
            Range("A1").Select
            ActiveWorkbook.Close True   '關閉並存檔
        End If
    Next
    '--------貼到 IPQC FQC 檢驗紀錄表-------
    
    
    '--------貼到首件檢驗紀錄表-------
    For Each ws_1 In Worksheets
        If LCase(ws_1.Name) = LCase(my_ws4) Then   '判斷是否已存在工作表，已存在直接複製貼上
            Workbooks("test首件檢驗紀錄表(組立).xlsx").Worksheets("員工名冊").Activate   '指定當前活頁簿、工作表
            Worksheets("員工名冊").Range("I1", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有工站資料
            
            Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("員工名冊").Activate   '指定當前活頁簿、工作表
            ActiveSheet.Range("I1", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).Select   '選取資料
            Selection.Copy  '複製
            
            Workbooks("test首件檢驗紀錄表(組立).xlsx").Worksheets("員工名冊").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("I1", ActiveSheet.Range("J" & ActiveSheet.Rows.Count).End(xlUp)).Select   '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            ' 排序 I 欄至 J 欄的資料
            ' Key1:=Range("I1")     依據 I 欄排序
            ' Order1:=xlAscending  升冪排序
            ' Header:=xlYes         有標題列
            Columns("I:J").Sort Key1:=Range("I1"), Order1:=xlAscending, Header:=xlYes  '依照編號排序
            
            Range("A1").Select
            ActiveWorkbook.Close True   '關閉並存檔
        End If
    Next
    '--------貼到首件檢驗紀錄表-------
End Sub




