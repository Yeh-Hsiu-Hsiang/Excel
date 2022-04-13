Sub 匯出TEST銷貨()

    Dim wb As String

    wb = ActiveWorkbook.Name
    
    Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\巨集\雅琪個人巨集\銷貨\TEST銷貨.xls"  '開啟檔案
    'Workbooks.Open Filename:="C:\Users\ywqa011\Desktop\雅琪\銷貨\TEST銷貨.xls"  '開啟檔案
    
    Range("A2", "AX" & Range("A65536").End(xlUp).Row + 1).Select
    Selection.Delete
    
    Workbooks(wb).Worksheets("轉銷貨單據欄位").Activate
    
    Range("A2", "AX" & Range("A65536").End(xlUp).Row).Select
    Selection.Copy
    
    Workbooks("TEST銷貨.xls").Worksheets(1).Activate
    Range("A2").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
End Sub
