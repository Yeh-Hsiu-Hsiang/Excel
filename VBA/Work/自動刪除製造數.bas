Sub �۰ʧR���s�y��()

    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

    For i = 6 To ActiveSheet.Range("R65536").End(xlUp).Row
        
        If Range("R" & i) = "" Then
            Rows(i).Select
            Selection.Delete Shift:=xlUp
        End If
    Next
    
End Sub
