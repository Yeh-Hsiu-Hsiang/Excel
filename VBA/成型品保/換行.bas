Attribute VB_Name = "����"
Sub ����()

    For i = 15 To Cells(1, Columns.Count).End(xlToLeft).Column  '�qO���̫�@��
    
        For j = 2 To Range("B65536").End(xlUp).Row  '��̫�@�C
        
            rawData = Cells(j, i)   ' ���o��l���

            If rawData <> "" Then
                fieldArray = Split(rawData, "#")    ' �ϥ� Split �������
                If UBound(fieldArray) <> 0 Then   '-1�Ű}�C
                    Cells(j, i) = fieldArray(0) & vbCrLf & "#" & fieldArray(1)   ' �N�U������J�������x�s��
                Else
                    GoTo ContinueForLoop
                End If
            Else
                GoTo ContinueForLoop
            End If
        
ContinueForLoop:
        Next j
        
    Next i
    
End Sub

