Attribute VB_Name = "加工每日入庫"
Sub 加工每日入庫()

    ActWb = ActiveWorkbook.Name
    
    '-----------製令日期-----------
    Range("L2").Select
    ActiveCell.Formula = "=TEXT(B2, ""m月d日"")"
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

    Selection.AutoFill Destination:=Range("L2:L" & lrow)
    Range("L2:L" & lrow).Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    '-----------製令日期-----------
    
    ActiveSheet.Range("A:K").AutoFilter Field:=7, Criteria1:="1"    '篩選欄號為1的資料
    Cells.Replace What:=".0000", Replacement:="", LookAt:=xlPart    '小數點後四位取代
    Range("E:F, H:K").Select    '把不需要的欄位隱藏
    Selection.EntireColumn.Hidden = True
    
    '-----------移動欄位位置-----------
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("D:D").Select
    Selection.Cut
    Range("A1").Select
    ActiveSheet.Paste
    Columns("E:E").Select
    Selection.Cut
    Range("D1").Select
    ActiveSheet.Paste
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    '-----------移動欄位位置-----------
    
    Range("A2", Range("D65536").End(xlUp)).Select
    Selection.Copy

    Workbooks.Open Filename:="\\yeawen\files-server\10_公用\00_i-Reporter 行動表單系統\碧垂\0617_加工每日入庫.xls"  '開啟檔案
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Name Like "*年" Then
            ws.Activate
            Exit For
        End If
    Next ws

    i = 3
    Do While True
        If ActiveSheet.Cells(i, 1).Value = "" Then
            ActiveSheet.Cells(i, 1).Select
            Exit Do
        End If
        i = i + 1
    Loop
    
    ActiveSheet.Select
    Selection.PasteSpecial xlPasteValues
    
End Sub
