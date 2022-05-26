
Sub 客戶訂單0轉O()

    For i = 2 To Range("E65536").End(xlUp).Row
    
        If Not IsNumeric(Range("K" & i)) And Left(Range("K" & i), 1) = "0" Then
            Range("K" & i).Replace "0", "O", xlPart
        End If
    Next
    
    MsgBox "轉換完成"

End Sub
