Attribute VB_Name = "�ƻs�W�@��"
Sub �ƻs�W�@��()
Attribute �ƻs�W�@��.VB_ProcData.VB_Invoke_Func = "w\n14"

' �ֳt��: Ctrl+w

    ActiveCell.Offset(-1, 0).Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
End Sub
