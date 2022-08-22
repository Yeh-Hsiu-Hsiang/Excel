Sub 取消換行()

    Cells.Replace _
    [vbLf], [""], xlPart
    
    Cells.Replace _
    [vbCr], [""], xlPart
    
    Cells.Replace _
    [vbCrLf], [""], xlPart
    
    Cells.Replace _
    [vbNewLine], [""], xlPart

End Sub
