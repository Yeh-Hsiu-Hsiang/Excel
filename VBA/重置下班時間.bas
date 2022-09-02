
Sub 重置下班時間()
Attribute 重置下班時間.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 重置下班時間 巨集
'
    Range("D5:D35").Select
    ActiveWindow.SmallScroll Down:=-15
    Selection.ClearContents
    
    Range("I5:I35").Select
    Selection.ClearContents
End Sub
