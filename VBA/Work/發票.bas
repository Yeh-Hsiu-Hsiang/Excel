Attribute VB_Name = "祇布"
Sub Bill()

    Dim ws, ws_2 As Worksheet, i, j As Integer, Find_cell As Range
    
    For Each ws In Workbooks(2).Worksheets
        For Each ws_2 In Workbooks(3).Worksheets

            If LCase(ws.Name) = LCase(ws_2.Name) Then   '耞琌
                
                'For i = 3 To Workbooks(2).Worksheets(ws.Name).Range("B65536").End(xlUp).Row
                    For j = 3 To Workbooks(3).Worksheets(ws_2.Name).Range("B65536").End(xlUp).Row
                        
                        '  A:J 絛瞅い碝т戈才纗
                        Set Find_cell_B = Workbooks(2).Worksheets(ws.Name).Range("A3:J65536").Find(What:=Workbooks(3).Worksheets(ws_2.Name).Range("B" & j), LookIn:=xlValues, LookAt:=xlWhole)
                        Set Find_cell_G = Workbooks(2).Worksheets(ws.Name).Range("A3:J65536").Find(What:=Workbooks(3).Worksheets(ws_2.Name).Range("G" & j), LookIn:=xlValues, LookAt:=xlWhole)
                        
                        ' 狦Τт
                        If Not Find_cell_B Is Nothing Then
                        
                            If Find_cell_B <> "" And Find_cell_B <> "祇布腹絏" Then
                                Workbooks(2).Worksheets(ws.Name).Activate
                                Range(Find_cell_B.Address).Select
                                
                                With Selection.Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 65535
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            End If
                        End If
                            
                        If Not Find_cell_G Is Nothing Then

                            If Find_cell_G <> "" And Find_cell_G <> "祇布腹絏" Then
                                Workbooks(2).Worksheets(ws.Name).Activate
                                Range(Find_cell_G.Address).Select

                                With Selection.Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 65535
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            End If
                        End If
                        
                        
                    Next j
                'Next i
            End If
        Next
    Next
    
End Sub
