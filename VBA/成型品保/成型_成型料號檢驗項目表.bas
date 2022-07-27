Attribute VB_Name = "成型_成型料號檢驗項目表"

Sub 成型_成型料號檢驗項目表()

Workbooks.Open fileName:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\03_設計檔案資料\品保課\成型品保\首件檢驗紀錄表(射出成型)_iPad.xlsx"  '開啟檔案
Workbooks.Open fileName:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\03_設計檔案資料\品保課\成型品保\成型射出_QC檢驗紀錄表_iPad.xlsx"  '開啟檔案


    Dim ws As Worksheet

    '--------貼到 IPQC FQC 檢驗紀錄表-------
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("成型料號檢驗項目表") Then   '判斷是否已存在工作表，已存在直接複製貼上
            
            Workbooks("成型射出_QC檢驗紀錄表_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '指定要上傳至iReporter檔案的活頁簿及工作表
            
            'Range("A1").SpecialCells(xlCellTypeLastCell)    最後一格有資料的位置
            Worksheets("成型料號檢驗項目表").Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).ClearContents  '清空舊有資料
            
            Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型料號檢驗項目表").Activate   '指定原本資料活頁簿、工作表
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select  '選取要複製的範圍
            Selection.Copy  '複製
            
            Workbooks("成型射出_QC檢驗紀錄表_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            ActiveWorkbook.Close True   '關閉並存檔
        End If
    Next
    '--------貼到 IPQC FQC 檢驗紀錄表-------
    
    

    '--------貼到首件檢驗紀錄表-------
    For Each ws_1 In Worksheets
        If LCase(ws_1.Name) = LCase("成型料號檢驗項目表") Then   '判斷是否已存在工作表，已存在直接複製貼上

            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '指定要上傳至iReporter檔案的活頁簿及工作表

            'Range("A1").SpecialCells(xlCellTypeLastCell)    最後一格有資料的位置
            Worksheets("成型料號檢驗項目表").Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).ClearContents  '清空舊有資料

            Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("成型料號檢驗項目表").Activate   '指定原本資料活頁簿、工作表
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select  '選取要複製的範圍
            Selection.Copy  '複製

            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("成型料號檢驗項目表").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("A1", Range("A1").SpecialCells(xlCellTypeLastCell)).Select '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            ActiveWorkbook.Close True   '關閉並存檔
        End If
    Next
    '--------貼到首件檢驗紀錄表-------
    
    Application.CutCopyMode = False

End Sub

