Sub �۰ʧR���s�y��()

    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate

    For i = 6 To ActiveSheet.Range("R65536").End(xlUp).Row
        
        If Range("R" & i) = "" Then
        
            Rows(i).Offset.Select
            Selection.Delete Shift:=xlUp
            
            If Range("R" & i).Offset(-1, 0) = "" Then
                Rows(i).Offset(-1, 0).Select
                Selection.Delete Shift:=xlUp
            End If
        End If
    Next
    
End Sub

