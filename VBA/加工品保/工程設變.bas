
Sub 工程設變()

    Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\設計檔案資料\品保課\加工品保\test加工組立_QC檢驗紀錄表_iPad.xlsx"  '開啟檔案

    Dim ws As Worksheet
    Dim my_ws1, my_ws2, my_ws3, my_ws4 As String
    
    my_ws1 = "加工QC檢驗項目表"
    my_ws2 = "生產異常狀況分析追蹤紀錄"
    my_ws3 = "工程設變"
    my_ws4 = "員工名冊"
    
    '工程設變
    For Each ws In Worksheets
        If LCase(ws.Name) = LCase(my_ws3) Then   '判斷是否已存在工作表，已存在直接複製貼上
        
            Workbooks("test加工組立_QC檢驗紀錄表_iPad.xlsx").Worksheets("工程設變").Activate   '指定當前活頁簿、工作表
            
            'Range("C1").SpecialCells(xlCellTypeLastCell)   到最後一格有資料的位置
            Worksheets("工程設變").Range("C1", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).ClearContents '清空舊有資料
            
            Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("工程設變").Activate   '指定當前活頁簿、工作表
            ActiveSheet.Range("C1", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select   '選取資料
            Selection.Copy  '複製
            
            Workbooks("test加工組立_QC檢驗紀錄表_iPad.xlsx").Worksheets("工程設變").Activate   '選擇要貼上的活頁簿、工作表
            ActiveSheet.Range("C1", ActiveSheet.Range("O" & ActiveSheet.Rows.Count).End(xlUp)).Select   '選取要貼上的範圍
            Selection.PasteSpecial  '貼上
            
            ' 排序 C 欄至 O 欄的資料
            ' Key1:=Range("D1")     依據 D 欄排序
            ' Order1:=xlDescending  降冪排序
            ' Header:=xlYes         有標題列
            Columns("C:O").Sort Key1:=Range("D1"), Order1:=xlDescending, Header:=xlYes  '依照日期排序
            
            '---------處理LM欄
            Range("L2").Select
            Selection.Formula = "=IF($K2="""","""", CONCATENATE($K2,COUNTIF($K$1:$K2,$K2)))"  '設定 L2儲存格公式
            Range("L2").Select  '選取L2
            Selection.Copy  '複製 L2公式
            
            Dim x As Integer
            x = Range("K65536").End(xlUp).Row   '根據K欄最後一筆資料來找資料共幾列
            
            Range("L2", "L" & x).Select
            Selection.PasteSpecial  '貼上公式
            
            Range("L2", "L" & x).Select
            Selection.Copy
            Selection.PasteSpecial xlPasteValues '只貼上值
            
            Range("M2").Select
            Selection.Formula = "=CONCATENATE(TEXT($D2,""YYYYMMDD""),""，"",$E2,""，"",$O2)"   '設定 M2儲存格值
            Range("M2").Select
            Selection.Copy
            
            Range("M2", "M" & x).Select
            Selection.PasteSpecial
            
            Range("M2", "M" & x).Select
            Selection.Copy
            Selection.PasteSpecial xlPasteValues '只貼上值
            '---------處理LM欄
            
            Range("C2").Select
            
            ActiveWorkbook.Close True   '關閉並存檔
        End If
    Next
    
End Sub


