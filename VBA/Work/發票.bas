Attribute VB_Name = "�o��"
Sub Bill()

    Dim ws, ws_2 As Worksheet, i, j As Integer, Find_cell As Range
    
    For Each ws In Workbooks(2).Worksheets
        For Each ws_2 In Workbooks(3).Worksheets

            If LCase(ws.Name) = LCase(ws_2.Name) Then   '�P�_�u�@��O�_�ۦP
                
                'For i = 3 To Workbooks(2).Worksheets(ws.Name).Range("B65536").End(xlUp).Row
                    For j = 3 To Workbooks(3).Worksheets(ws_2.Name).Range("B65536").End(xlUp).Row
                        
                        ' �b A:J �d�򤤡A�M���ƲŦX���x�s��
                        Set Find_cell_B = Workbooks(2).Worksheets(ws.Name).Range("A3:J65536").Find(What:=Workbooks(3).Worksheets(ws_2.Name).Range("B" & j), LookIn:=xlValues, LookAt:=xlWhole)
                        Set Find_cell_G = Workbooks(2).Worksheets(ws.Name).Range("A3:J65536").Find(What:=Workbooks(3).Worksheets(ws_2.Name).Range("G" & j), LookIn:=xlValues, LookAt:=xlWhole)
                        
                        ' �p�G�����
                        If Not Find_cell_B Is Nothing Then
                        
                            If Find_cell_B <> "" And Find_cell_B <> "�o�����X" Then
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

                            If Find_cell_G <> "" And Find_cell_G <> "�o�����X" Then
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
