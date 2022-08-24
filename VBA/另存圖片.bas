Sub 另存圖片()

    Dim shp As Shape
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        If .Show = -1 Then
            pathn = .SelectedItems(1)
            
            For Each Item In ActiveSheet.Shapes
                'a = Item.Name & ".jpg"
                Set Rng = Item.TopLeftCell.Offset(0, -1)
                a = Range(Rng.Address).Value & ".jpg"
                Item.CopyPicture
    
                With ActiveSheet.ChartObjects.Add(0, 0, Item.Width, Item.Height).Chart
                    .ChartArea.Select
                    .Paste
                    .Export pathn & "\" & a
                    .Parent.Delete
                End With
            Next Item
        End If
    End With
   
End Sub



