Attribute VB_Name = "���ƶ�����"
Sub ���ƶ�����()

    For j = 2 To Range("B65536").End(xlUp).Row

        If Range("B" & j) = Range("B" & j).Offset(-1, 0) And Range("B" & j) <> "" Then
            Rows(j).Select
            Selection.Delete Shift:=xlUp
            j = j - 1
        End If
    Next


End Sub
