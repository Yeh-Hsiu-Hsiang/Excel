Sub 產品編號版本更新()

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1") = "版本"

    Application.DisplayAlerts = False
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(12, 1)), TrailingMinusNumbers:=True
    Application.DisplayAlerts = True

    Range("H1") = "版本排序"
    Range("H2").Select
    ActiveCell.Formula = _
        "=IF(IFERROR(IF(FIND(""-"",B2, 1)>=1,"""",B2),B2)=0,"""",IFERROR(IF(FIND(""-"",B2, 1)>=1,"""",B2),B2))"


    Range("I1") = "產品編號"
    Range("I2").Select
    ActiveCell.Formula = "=A2 & B2"
    
    Range("J1") = "停用"
    Range("J2").Select
    ActiveCell.Formula = "=IF(G2=""  /  /  "","""",G2)"

    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & lrow)

    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I" & lrow)
    
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J" & lrow)

    For i = 1 To Range("H65536").End(xlUp).Row
        If Range("H" & i) = "" Then
            Rows(i).Select
            Selection.Delete Shift:=xlUp
        End If
    Next


    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "A2:A" & lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "H2:H" & lrow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A2:H" & lrow)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


    For j = 2 To Range("A65536").End(xlUp).Row

        If Range("A" & j) = Range("A" & j).Offset(-1, 0) And Range("A" & j) <> "" Then
            Rows(j).Select
            Selection.Delete Shift:=xlUp
            j = j - 1
        End If
        
        If Range("J" & j) <> "" Then
            Rows(j).Select
            Selection.Delete Shift:=xlUp
            j = j - 1
        End If
    Next

    Range("I:I").Copy

    Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

    Range("A:A").PasteSpecial xlPasteValues

    Worksheets(1).Activate
    Range("D:E").Copy

    Worksheets(2).Activate
    Range("B:C").PasteSpecial xlPasteValues
    
    [B:B].Select
    With Selection
        .NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* "" - ""??_-;_-@_-"
        .Value = .Value
    End With
    
End Sub
