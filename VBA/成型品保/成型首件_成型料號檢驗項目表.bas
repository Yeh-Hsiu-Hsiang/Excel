
Sub 成型首件_成型料號檢驗項目表()

Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\03_設計檔案資料\品保課\成型品保\首件檢驗紀錄表(射出成型)_iPad.xlsx"  '開啟檔案

    Dim ws As Worksheet

    '成型料號檢驗項目表
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("成型料號檢驗項目表") Then   '判斷是否已存在工作表，已存在直接複製貼上
        
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '指定要上傳至iReporter檔案的活頁簿及工作表
            
            Worksheets("成型料號檢驗項目表").Range("B1", ActiveSheet.Range("AT" & Range("B65536").End(xlUp).Row)).ClearContents '清空舊有資料
            
            
            '複製母件編號、子件料號、品名規格、客戶、Fa機台、順序號、模號
            Workbooks("20210330.xlsm").Worksheets("成型料號檢驗項目表").Activate   '指定原本資料活頁簿、工作表
            ActiveSheet.Range("B3", ActiveSheet.Range("H" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("B1", ActiveSheet.Range("H" & Range("B65536").End(xlUp).Row)).Select    '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            '複製穴數、週期
            Workbooks("20210330.xlsm").Worksheets("成型料號檢驗項目表").Activate
            ActiveSheet.Range("L3", ActiveSheet.Range("M" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("I1", ActiveSheet.Range("J" & Range("B65536").End(xlUp).Row)).Select    '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            '複製原料
            Workbooks("20210330.xlsm").Worksheets("成型料號檢驗項目表").Activate
            ActiveSheet.Range("I3", ActiveSheet.Range("I" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("K1", ActiveSheet.Range("K" & Range("B65536").End(xlUp).Row)).Select    '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            
            '複製 SOP、SIP、標準樣
            Workbooks("20210330.xlsm").Worksheets("成型料號檢驗項目表").Activate
            ActiveSheet.Range("X3", ActiveSheet.Range("Z" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("L1", ActiveSheet.Range("N" & Range("B65536").End(xlUp).Row)).Select    '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            
            '複製 檢驗項目
            Workbooks("20210330.xlsm").Worksheets("成型料號檢驗項目表").Activate
            ActiveSheet.Range("AO3", ActiveSheet.Range("BT" & Range("B65536").End(xlUp).Row)).Select
            Selection.Copy
            
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("O1", ActiveSheet.Range("AT" & Range("B65536").End(xlUp).Row)).Select    '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            
            
            Range("B1").Select
            Application.CutCopyMode = False
            ActiveWorkbook.Close True   '關閉並存檔
        End If
    Next
End Sub

