Attribute VB_Name = "成型_工程設變"
Sub 成型_工程設變()

    Workbooks.Open fileName:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\03_設計檔案資料\品保課\成型品保\首件檢驗紀錄表(射出成型)_iPad.xlsx"  '開啟檔案
    Workbooks.Open fileName:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\03_設計檔案資料\品保課\成型品保\成型射出_QC檢驗紀錄表_iPad.xlsx"  '開啟檔案
    Workbooks.Open fileName:="\\yeawen\files-server\05_品保\13-3樓組立(儷秋)\品保IPQC_FQC日報系統(組立20210305.xlsm"

    Dim ws As Worksheet

    '工程設變
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("工程設變") Then   '判斷是否已存在工作表，已存在直接複製貼上
        
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("工程設變").Activate   '指定首件的活頁簿及工作表
            Worksheets("工程設變").Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).ClearContents '清空舊有資料
            
            Workbooks("成型射出_QC檢驗紀錄表_iPad.xlsx").Worksheets("工程設變").Activate   '指定QC檢驗紀錄表的活頁簿及工作表
            Worksheets("工程設變").Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).ClearContents '清空舊有資料
            
            Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("工程設變").Activate   '指定原本資料活頁簿、工作表
            ActiveSheet.Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).Select   '選取資料
            Selection.Copy  '複製
            
            Workbooks("品保IPQC_FQC日報系統(成型).xlsm").Worksheets("工程設變").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).Select   '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("工程設變").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).Select   '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            Workbooks("成型射出_QC檢驗紀錄表_iPad.xlsx").Worksheets("工程設變").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("C1", ActiveSheet.Range("P" & Range("K65536").End(xlUp).Row)).Select   '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            
            '-------------------------成型射出_QC檢驗紀錄表-------------------------
            
            ' 排序 C 欄至 P 欄的資料
            ' Key1:=Range("D1")     依據 D 欄排序
            ' Order1:=xlDescending  降冪排序
            ' Header:=xlYes         有標題列
            Columns("C:P").Sort Key1:=Range("D1"), Order1:=xlDescending, Header:=xlYes  '依照日期排序
            
'            '---------處理LM欄
'            Range("L2").Select
'            Selection.Formula = "=IF($K2="""","""", CONCATENATE($K2,COUNTIF($K$1:$K2,$K2)))"  '設定 L2儲存格公式
'            Range("L2").Select  '選取L2
'            Selection.Copy  '複製 L2公式
'
'            Dim x As Integer
'            x = Range("D65536").End(xlUp).Row   '根據 D 欄最後一筆資料來找資料共幾列
'
'            Range("L2", "L" & x).Select
'            Selection.PasteSpecial  '貼上公式
'
'            Range("L2", "L" & x).Select
'            Selection.Copy
'            Selection.PasteSpecial xlPasteValues '只貼上值
'
'            Range("M2").Select
'            Selection.Formula = "=CONCATENATE(TEXT($D2,""YYYYMMDD""),""，"",$E2,""，"",$O2)"   '設定 M2儲存格值
'            Range("M2").Select
'            Selection.Copy
'
'            Range("M2", "M" & x).Select
'            Selection.PasteSpecial
'
'            Range("M2", "M" & x).Select
'            Selection.Copy
'            Selection.PasteSpecial xlPasteValues '只貼上值
            '---------處理LM欄
            
            Range("D2").Select
            
            '-------------------------成型射出_QC檢驗紀錄表-------------------------
            
            
            '-------------------------首件檢驗紀錄表-------------------------
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("工程設變").Activate   '選擇活頁簿、工作表
            
            Columns("C:P").Sort Key1:=Range("D1"), Order1:=xlDescending, Header:=xlYes  '依照日期排序
            
'            '---------處理LM欄
'            Range("L2").Select
'            Selection.Formula = "=IF($K2="""","""", CONCATENATE($K2,COUNTIF($K$1:$K2,$K2)))"  '設定 L2儲存格公式
'            Range("L2").Select  '選取L2
'            Selection.Copy  '複製 L2公式
'
'            Dim y As Integer
'            y = Range("D65536").End(xlUp).Row   '根據 D 欄最後一筆資料來找資料共幾列
'
'            Range("L2", "L" & y).Select
'            Selection.PasteSpecial  '貼上公式
'
'            Range("L2", "L" & y).Select
'            Selection.Copy
'            Selection.PasteSpecial xlPasteValues '只貼上值
'
'            Range("M2").Select
'            Selection.Formula = "=CONCATENATE(TEXT($D2,""YYYYMMDD""),""，"",$E2,""，"",$O2)"   '設定 M2儲存格值
'            Range("M2").Select
'            Selection.Copy
'
'            Range("M2", "M" & y).Select
'            Selection.PasteSpecial
'
'            Range("M2", "M" & y).Select
'            Selection.Copy
'            Selection.PasteSpecial xlPasteValues '只貼上值
'            '---------處理LM欄
            
            Range("D2").Select
            
            '-------------------------首件檢驗紀錄表-------------------------
            
            Application.CutCopyMode = False
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Close True   '關閉並存檔
            Workbooks("成型射出_QC檢驗紀錄表_iPad.xlsx").Close True   '關閉並存檔
            Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Close False
        End If
    Next
    
End Sub



