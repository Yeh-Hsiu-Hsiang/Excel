Attribute VB_Name = "BOM_��s��Ϯ榡"
Sub BOM_��s��Ϯ榡()
Attribute BOM_��s��Ϯ榡.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim ActWb As String
    ActWb = ActiveWorkbook.Name

    Range("A:DZ").UnMerge

    For i = 1 To 1000
        If InStr(1, Cells(3, i), "�s���") = 1 Then
            Columns(i).Offset(0, 1).Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(3, i).Offset(0, 1) = "�������"
            
            For j = 5 To 1000 Step 2
                If Cells(j, i) <> "" Then
                    Cells(j, i).Select
                    Selection.Cut Destination:=Cells(j, i).Offset(-1, 1)
                End If
            Next
        End If
        
        If InStr(1, Cells(3, i), "�s��") = 1 Then
            Columns(i).Offset(0, 1).Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(3, i).Offset(0, 1) = "�������"
            
            For j = 5 To 700 Step 2
                If Cells(j, i) <> "" Then
                    Cells(j, i).Select
                    Selection.Cut Destination:=Cells(j, i).Offset(-1, 1)
                End If
            Next
        End If
        
        If Cells(2, i) <> "" And Cells(3, i) = "" Then
            Cells(2, i).Resize(2, 1).Select
            Selection.Merge
        End If
        
        If Cells(2, i) <> "" And Cells(2, i).Offset(0, 1) = "" Then
            Cells(2, i).Resize(1, 2).Select
            Selection.Merge
        End If
    Next


    For k = 5 To 1000
        If Range("A" & k) = "" Then
            Rows(k).Select
            Selection.Delete Shift:=xlUp
        End If
    Next

    Range("A4").Select
    
End Sub
