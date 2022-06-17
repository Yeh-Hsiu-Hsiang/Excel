Attribute VB_Name = "新增"
Sub 新增()

    Application.ScreenUpdating = False

    If Sheets("輸入").Range("E2") = "" Then
        
        Sheets("輸入").Range("AI4:HO4").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("客戶主檔").Select
        
        Range("A3").Select
        
        '----------讓選定的位置為有資料的最底行----------
        i = 3
        Do While True
            If ActiveSheet.Cells(i, 1).Value = "" Then
                ActiveSheet.Cells(i, 1).Select
                Exit Do
            End If
            i = i + 1
        Loop
        '----------讓選定的位置為有資料的最底行----------
        
        Selection.PasteSpecial Paste:=xlPasteValues
        
        Range("A3").Select
        
        Sheets("輸入").Select
        Range("E2").Select
        
        清除客戶表資料

    Else
    
        DialogSheets("錯誤訊息視窗").Show
    
        ' MsgBox "資料儲存方式選擇錯誤, 請再確認！ 如果確定要新增為新客戶，  請清除 E2  儲存格資料後，再按下新增　　　。"
    End If

End Sub

Sub 清除客戶表資料()

    Sheets("輸入").Select
    Range("E2, E5, I5, D8, F8, J8, L8, S8, Y8, E11, H11, K11, P11").Select
    Application.CutCopyMode = False
    Selection.ClearContents

    '--------階層次序--------
    For i = 28 To ActiveSheet.Range("C65536").End(xlUp).Row + 1 Step 2
        Range("C" & i & ":Z" & i).Select
        Selection.ClearContents
    Next i
    '--------階層次序--------
    
    '--------版本--------
    For j = 15 To 23 Step 2
        Range("C" & j & ":Z" & j).Select
        Selection.ClearContents
    Next j
    '--------版本--------
    
    
    '------------刪除BOM、成品圖、FA------------
    Range("D117:F117").Select
    Selection.ClearContents
    '------------刪除BOM、成品圖、FA------------
    
    
    '------------刪除零件圖------------
    For k = 4 To 16
        '------------刪除零件圖1~10------------
        Cells(121, k).Select
        Selection.ClearContents
        '------------刪除零件圖1~10------------
        
        
        '------------刪除日期版本1~10------------
        Cells(124, k).Select
        Selection.ClearContents
        '------------刪除日期版本1~10------------


        '------------刪除零件圖11~20------------
        Cells(128, k).Select
        Selection.ClearContents
        '------------刪除零件圖11~20------------


        '------------刪除日期版本11~20------------
        Cells(131, k).Select
        Selection.ClearContents
        '------------刪除日期版本11~20------------


        '------------刪除零件圖21~30------------
        Cells(135, k).Select
        Selection.ClearContents
        '------------刪除零件圖21~30------------


        '------------刪除日期版本21~30------------
        Cells(138, k).Select
        Selection.ClearContents
        '------------刪除日期版本21~30------------
    Next
    '------------刪除零件圖------------
    
    
    
    '------------刪除成品------------
    Range("D144").Select
    Selection.ClearContents
    '------------刪除成品------------
    
    
    '------------刪除零件------------
    For l = 4 To 16
        '------------刪除零件1~10------------
        Cells(148, l).Select
        Selection.ClearContents
        '------------刪除零件1~10------------
        
        
        '------------刪除日期版本1~10------------
        Cells(151, k).Select
        Selection.ClearContents
        '------------刪除日期版本1~10------------


        '------------刪除零件11~20------------
        Cells(155, k).Select
        Selection.ClearContents
        '------------刪除零件11~20------------


        '------------刪除日期版本11~20------------
        Cells(158, k).Select
        Selection.ClearContents
        '------------刪除日期版本11~20------------


        '------------刪除零件21~30------------
        Cells(162, k).Select
        Selection.ClearContents
        '------------刪除零件21~30------------


        '------------刪除日期版本21~30------------
        Cells(165, k).Select
        Selection.ClearContents
        '------------刪除日期版本21~30------------
    Next
    '------------刪除零件------------
    
    Range("E2").Select
End Sub



