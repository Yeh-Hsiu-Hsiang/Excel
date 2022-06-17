Attribute VB_Name = "LTA庫存轉換"
Sub LTA庫存轉換()

    Dim ActWb As String

    ActWb = ActiveWorkbook.Name
    
    '------------------讀取並開啟最新每日庫存------------------
    Dim pth As String, fn As String, ary(), tmpMax As Long, n As Integer, wb As Workbook

    pth = "\\yeawen\files-server\06_資材\01_生管\航電每日資訊\每日庫存\"    '設置路徑
    'pth = "C:\Users\ywqa011\Desktop\每日庫存\"    '設置路徑
    
    fn = Dir(pth & "*.xls")     '瀏覽資料夾下的 .xls文件
    n = 0: tmpMax = 0
    
    Do While fn <> ""
        If fn <> ThisWorkbook.Name Then
            n = n + 1
            ReDim Preserve ary(n)
            ary(n) = Left(Right(fn, 11), 7)   '放入 excel 檔名的日期
            If ary(n) > tmpMax Then
                tmpMax = ary(n)   '最新日期檔案
            End If
        End If
        fn = Dir
    Loop
    
    Set wb = Workbooks.Open(pth & "MERP每日庫存" & tmpMax & ".xls", , True)
    '------------------讀取並開啟最新每日庫存------------------
    
    
    Workbooks(ActWb).Worksheets("LTA").Activate

    Dim i, j, k As Integer, Find_Value As Long
    
    For k = 8 To Workbooks(ActWb).Worksheets("LTA").Cells(2, Columns.Count).End(xlToLeft).Column '最後一欄
        If InStr(1, Cells(2, k), Format(Date, "MM/DD")) = 1 Then    '判斷是否等於今天
            For i = 3 To Workbooks(ActWb).Worksheets("LTA").Range("C65536").End(xlUp).Row - 1
                For j = 5 To wb.Worksheets("產品存量").Range("A65536").End(xlUp).Row
        
                    If Left(wb.Worksheets("產品存量").Range("A" & j), 12) = Workbooks(ActWb).Worksheets("LTA").Range("C" & i) Then  '判斷 LTA 料號等同於本公司料號
                        Find_Value = Find_Value + wb.Worksheets("產品存量").Range("A" & j).Offset(0, 2)   '料號庫存數加總
                    End If
                Next j
                
                Workbooks(ActWb).Worksheets("LTA").Activate
                Workbooks(ActWb).Worksheets("LTA").Cells(i, k).Value = Find_Value
                
                Find_Value = 0
            Next i
        End If
    Next k
     
    '------------------條件式格式設定------------------
    Range("A3:AE9").Select
    Range(Selection, Selection.End(xlDown)).FormatConditions.Delete '清除格式
    Range(Selection, Selection.End(xlDown)).Select  '選取範圍
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=OR(ISNUMBER(SEARCH(""-"", $AH9, 1)), ISNUMBER(SEARCH(""-"", $AI9, 1)), ISNUMBER(SEARCH(""-"", $AJ9, 1)))" '設定條件公式
    
    With Selection.FormatConditions(1).Interior '設定格式
        .PatternColorIndex = xlAutomatic
        .Color = 10066431
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    '------------------條件式格式設定------------------
    
End Sub

