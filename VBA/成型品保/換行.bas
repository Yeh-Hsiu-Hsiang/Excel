Attribute VB_Name = "換行"
Sub 換行()

    For i = 15 To Cells(1, Columns.Count).End(xlToLeft).Column  '從O欄到最後一欄
    
        For j = 2 To Range("B65536").End(xlUp).Row  '到最後一列
        
            rawData = Cells(j, i)   ' 取得原始資料

            If rawData <> "" Then
                fieldArray = Split(rawData, "#")    ' 使用 Split 分割欄位
                If UBound(fieldArray) <> 0 Then   '-1空陣列
                    Cells(j, i) = fieldArray(0) & vbCrLf & "#" & fieldArray(1)   ' 將各個欄位填入對應的儲存格
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

