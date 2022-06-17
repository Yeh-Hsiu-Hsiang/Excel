Attribute VB_Name = "轉當日資料"

Sub 轉當日資料()
Attribute 轉當日資料.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 轉當日資料 巨集
'
' 快速鍵: Ctrl+w
'
Range("A2:C2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("A7:BH7").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=6
    Selection.Copy
       
    
    Sheets("全月報表").Select
    
    INN7
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    '--------- 轉全月資料到總表AF ---------
    Workbooks.Open Filename:="\\yeawen\files-server\02_成型\成型每日資訊\成型日報表\雅文_成型生產總報表_AF.xlsm"
    
    Sheets("全月份報表").Select
    
    INN7
    
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("G:R, V:AA, AE:AM, AO:BH").ClearContents   '清除資料內容
    '--------- 轉全月資料到總表AF ---------
        
    
    
    MsgBox "資料轉換完成!"
    
End Sub
Sub INN7()

'讓選定的位置為有資料的最底行

i = 7
    Do While True
        If ActiveSheet.Cells(i, 1).Value = "" Then
            ActiveSheet.Cells(i, 1).Select
            Exit Do
        End If
        i = i + 1
    Loop
  
    
End Sub
Sub 隱藏0列()

Application.ScreenUpdating = False

ASD = ActiveCell.Column


Rows("78:105").Select

Selection.EntireRow.Hidden = False

For i = 78 To 105

If Cells(i, ASD) = "0" Then

Rows(i).Hidden = True

End If

Next

End Sub

