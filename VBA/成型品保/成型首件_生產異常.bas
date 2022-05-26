
Sub 成型首件_生產異常狀況分析追蹤紀錄()

    Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\03_設計檔案資料\品保課\成型品保\首件檢驗紀錄表(射出成型)_iPad.xlsx"  '開啟檔案

    Dim ws As Worksheet

    '生產異常狀況分析追蹤紀錄
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase("生產異常狀況分析追蹤紀錄") Then   '判斷是否已存在工作表，已存在直接複製貼上
        
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("生產異常狀況分析追蹤紀錄").Activate   '指定當前活頁簿、工作表
            
            Worksheets("生產異常狀況分析追蹤紀錄").Range("A1", ActiveSheet.Range("AJ" & Range("A65536").End(xlUp).Row)).ClearContents '清空舊有資料
            
            Workbooks("20210330.xlsm").Worksheets("生產異常狀況分析追蹤紀錄").Activate   '指定當前活頁簿、工作表
            ActiveSheet.Range("A1", ActiveSheet.Range("AJ" & Range("D65536").End(xlUp).Row)).Select   '選取要複製的範圍
            Selection.Copy  '複製
            
            Workbooks("首件檢驗紀錄表(射出成型)_iPad.xlsx").Worksheets("生產異常狀況分析追蹤紀錄").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("A1", ActiveSheet.Range("AJ" & Range("D65536").End(xlUp).Row)).Select   '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            ' 排序 D 欄至 L 欄的資料
            ' Key1:=Range("E1")     依據 E 欄排序
            ' Order1:=xlDescending  降冪排序
            ' Header:=xlYes         有標題列
            Columns("A:AJ").Sort Key1:=Range("E1"), Order1:=xlDescending, Header:=xlYes  '依照日期排序
            
            
            '---------處理AB欄
            Range("A2").Select
            Selection.Formula = "=IF($D2="""","""", D2&COUNTIF($D$1:$D2,$D2))"  '設定 A2儲存格公式
            Range("A2").Select  '選取A2
            Selection.Copy  '複製 A2公式
            
            Dim x As Integer
            x = Range("D65536").End(xlUp).Row   '根據D欄最後一筆資料來找資料共幾列
            
            Range("A2", "A" & x).Select
            Selection.PasteSpecial  '貼上公式
            
            Range("A2", "A" & x).Select
            Selection.Copy
            Selection.PasteSpecial xlPasteValues '只貼上值
            
            Range("B2").Select
            Selection.Formula = "=CONCATENATE(TEXT($E2,""YYYYMMDD""),""，"",$H2)"   '設定 B2儲存格值
            Range("B2").Select
            Selection.Copy
            
            Range("B2", "B" & x).Select
            Selection.PasteSpecial
            
            Range("B2", "B" & x).Select
            Selection.Copy
            Selection.PasteSpecial xlPasteValues '只貼上值
            '---------處理AB欄
            
            Range("A1").Select
            
            ActiveWorkbook.Close True   '關閉並存檔
        End If
    Next


End Sub
