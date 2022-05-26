
Sub MARS訂單轉銷貨單據()


'
' 巨集7 巨集
'
' 快速鍵: Ctrl+u
'
                ' MARS訂單轉銷貨單據 巨集
'
'
Application.ScreenUpdating = False
      U = ActiveSheet.Name  'ActiveSheet.Name
    
   
    
   If U = ActiveSheet.Range("c3") & " " & ActiveSheet.Range("i3") Then
    
         
         資料轉轉
         Else
         
  End If
    Range("F11").Select
End Sub

Sub 資料轉轉()
'
' 資料轉轉 巨集
'
'Set U = n
 rn = ActiveSheet.Range("CL2").Value  '單據頁貼上的筆數列位
 en = ActiveSheet.Range("CL4").Value   '整理頁貼上的筆數起始列位
 yn = ActiveSheet.Range("CL3").Value   '整理頁貼上的筆數結束列位
'
U = ActiveSheet.Range("c3") & " " & ActiveSheet.Range("i3")

'-----------
    Range("AC7:AG7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("CM" & en).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
'-----------
'    Range("CM7").Select
'    Application.CutCopyMode = False
'    ActiveSheet.Range("$CM$6:$CQ$157").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5) _
'        , Header:=xlYes
'----------

    ActiveWorkbook.Worksheets(U).Sort.SortFields.Clear
    'ActiveWorkbook.Worksheets(U).Sort.SortFields.Add2 Key:=Range( _
        "CM7:CQ156"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
        
    With ActiveWorkbook.Worksheets(U).Sort
        .SetRange Range("CM7:CQ156")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .SortFields.Add Key:=Range("CO7"), Order:=xlDescending
        .Apply
    End With
        
'------------

    Selection.End(xlDown).Select
    Selection.ClearContents
    
    Range("CM7").Resize(Range("CL4").Value - 7, 1).Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("轉銷貨單據欄位").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("AR" & rn).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
   
    Sheets(U).Select
    Range("CO7").Resize(Range("CL4").Value - 7, 1).Select
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("轉銷貨單據欄位").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("AJ" & rn).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets(U).Select
    Range("CP7").Resize(Range("CL4").Value - 7, 1).Select
    
    'Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("轉銷貨單據欄位").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("AM" & rn).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("AN2").Select
    Sheets(U).Select
    Range("CQ7").Resize(Range("CL4").Value - 7, 1).Select
    
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("轉銷貨單據欄位").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("AN" & rn).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
     Sheets(U).Select
    ActiveWindow.SmallScroll Down:=-60
    Range("CN7").Resize(Range("CL4").Value - 7, 1).Select
    
    
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("轉銷貨單據欄位").Select
    Range("A" & rn).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    Sheets(U).Select
    Range("B6").Select
    
    Range("T3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "資料完成轉換"
    Selection.End(xlUp).Select
    Range("T3").Select
End Sub
