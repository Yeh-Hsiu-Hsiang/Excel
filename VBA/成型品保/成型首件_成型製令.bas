
Sub 成型首件_成型製令()

Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\03_設計檔案資料\品保課\成型品保\首件檢驗紀錄表(射出成型)_iPad.xlsx"  '開啟檔案

    Dim ws As Worksheet

    '成型製令
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("成型製令") Then   '判斷是否已存在工作表，已存在直接複製貼上
        
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型製令").Activate   '指定要上傳至iReporter檔案的活頁簿及工作表
            
            Worksheets("成型製令").Range("A1", ActiveSheet.Range("S" & Range("A65536").End(xlUp).Row)).ClearContents '清空舊有資料
            
            Workbooks("20210330.xlsm").Worksheets("成型製令").Activate   '指定原本資料活頁簿、工作表
            ActiveSheet.Range("A2", ActiveSheet.Range("S" & Range("A65536").End(xlUp).Row)).Select   '選取資料
            Selection.Copy  '複製
            
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型製令").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("A1", ActiveSheet.Range("S" & Range("A65536").End(xlUp).Row)).Select    '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            Range("A1").Select
            Application.CutCopyMode = False
            ActiveWorkbook.Close True   '關閉並存檔
        End If
    Next
End Sub

