
Sub �Ȥ�q��0��O()

    For i = 2 To Range("E65536").End(xlUp).Row
    
        If Not IsNumeric(Range("K" & i)) And Left(Range("K" & i), 1) = "0" Then
            Range("K" & i).Replace "0", "O", xlPart
        End If
    Next
    
    MsgBox "�ഫ����"

End Sub
