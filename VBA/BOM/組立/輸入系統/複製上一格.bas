Attribute VB_Name = "複製上一格"
Sub 複製上一格()
Attribute 複製上一格.VB_ProcData.VB_Invoke_Func = "w\n14"

' 快速鍵: Ctrl+w

    ActiveCell.Offset(-1, 0).Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
End Sub
