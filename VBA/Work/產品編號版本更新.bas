Sub 產品編號版本更新()

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1") = "版本"

    Application.DisplayAlerts = False
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(12, 1)), TrailingMinusNumbers:=True
    Application.DisplayAlerts = True

    Range("G1") = "版本排序"
    Range("G2").Select
    ActiveCell.Formula = _
        "=IF(IFERROR(IF(FIND(""-"",RC[-5], 1)>=1,"""",RC[-5]),RC[-5])=0,"""",IFERROR(IF(FIND(""-"",RC[-5], 1)>=1,"""",RC[-5]),RC[-5]))"


    Range("H1") = "產品編號"
    Range("H2").Select
    ActiveCell.Formula = "=A2 & B2"

    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & lrow)

    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & lrow)

    For i = 1 To Range("G65536").End(xlUp).Row
        If Range("G" & i) = "" Then
            Rows(i).Select
            Selection.Delete Shift:=xlUp
        End If
    Next


    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "A2:A" & lrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "G2:G" & lrow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A2:G" & lrow)
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
    Next

    Range("H:H").Copy

    Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

    Range("A:A").PasteSpecial xlPasteValues

    Worksheets(1).Activate
    Range("D:E").Copy

    Worksheets(2).Activate
    Range("B:B").PasteSpecial xlPasteValues
    
    [B:B].Select
    With Selection
        .NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* "" - ""??_-;_-@_-"
        .Value = .Value
    End With
    
End Sub
