Sub Input_GeneralTable()
    
    Dim ActWb As String, i, j, k As Long
    
    ActWb = ActiveWorkbook.Name
    
    For j = 2 To Workbooks(ActWb).Worksheets(1).Range("A65536").End(xlUp).Row '�̫�@�C
    
        '------------��w��m������ƪ��̩���------------
        Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
        
        i = 6
        Do While True
            If Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Cells(i, "D").Value = "" Then
                Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Cells(i, "D").Select
                Exit Do
            End If
            i = i + 1
        Loop
        '------------��w��m������ƪ��̩���------------
    
        
        
        '----------�P�_�`��D��O ����BIPQC----------
        If InStr(1, ActWb, "����") <> 0 Then    '�P�_�O�_������
            Range("D" & i) = "����"
        ElseIf InStr(1, ActWb, "QC") <> 0 Then    '�P�_�O�_��IPQC
            Range("D" & i) = "IPQC"
        End If
        '----------�P�_�`��D��O ����BIPQC----------


        Workbooks(ActWb).Worksheets(1).Activate
        For k = 1 To Workbooks(ActWb).Worksheets(1).Cells(1, Columns.Count).End(xlToLeft).Column '�̫�@��
        
            '----------������----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "QR Code_���") <> 0 Then   '�P�_�O�_������
                Cells(j, k).Select
                Selection.Copy
            
                Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                Range("E" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------������----------
            
            
            '----------�����----------
            Dim inspector As String
            
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "�����") <> 0 Or InStr(1, Cells(1, k), "�����") Then   '�P�_�O�_���������
                inspector = inspector & Cells(j, k) & Chr(13)
        
                Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                Range("F" & i) = inspector
            End If
            '----------�����----------
            

            '----------�u���----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "QR Code_�u���") <> 0 Then   '�P�_�O�_����u���
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                Range("G" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------�u���----------
            
            
            '----------�s�O�渹----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "QR Code_�s�O�u��") <> 0 Then   '�P�_�O�_����s�O�渹
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                Range("H" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------�s�O�渹----------
            
            
            '----------�Ȥ�----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "�Ȥ�") <> 0 Then   '�P�_�O�_����Ȥ�
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                Range("J" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------�Ȥ�----------
            
            
            '----------����----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "QR Code_�Ƹ�") <> 0 Then   '�P�_�O�_����Ƹ�
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                Range("K" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------����----------
            

            '----------�~�W----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "�~�W") <> 0 Then   '�P�_�O�_����~�W
                Cells(j, k).Select
                Selection.Copy
        
                Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                Range("L" & i).Select
                Selection.PasteSpecial xlPasteValues
            End If
            '----------�~�W----------
            
            '----------SOP----------
            Workbooks(ActWb).Worksheets(1).Activate
    
            If InStr(1, Cells(1, k), "�@�~�W�d_SOP") <> 0 Then   '�P�_�O�_����SOP
                If InStr(1, Cells(j, k), "�i") = 1 Then
                    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                    Range("M" & i) = "V"
                Else
                    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                    Range("M" & i) = "X"
                End If
            End If
            '----------SOP----------
            
            '----------SIP----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "�@�~�W�d_SIP") <> 0 Then   '�P�_�O�_����SIP
                If InStr(1, Range("J" & j), "�i") = 1 Then
                    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                    Range("N" & i) = "V"
                Else
                    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                    Range("N" & i) = "X"
                End If
            End If
            '----------SIP----------
            
            '----------�˫~----------
            Workbooks(ActWb).Worksheets(1).Activate
            
            If InStr(1, Cells(1, k), "�@�~�W�d_�˫~") <> 0 Then   '�P�_�O�_����˫~
                If InStr(1, Range("K" & j), "�i") = 1 Then
                    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                    Range("O" & i) = "V"
                Else
                    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                    Range("O" & i) = "X"
                End If
            End If
            '----------�˫~----------
                        
            
            '----------�s�y��----------
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
    
            If Range("D" & i) = "IPQC" Then    '�P�_�OIPQC�A����b���~�ˬd��
                Workbooks(ActWb).Worksheets(1).Activate
                
                If InStr(1, Cells(1, k), "�b���~�ˬd��") <> 0 Then   '�P�_�O�_����b���~�ˬd��
                
                    Dim Semi_Finished_Rroduct As Long
                    
                    Semi_Finished_Rroduct = Semi_Finished_Rroduct + Cells(j, k)

                    Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                    Range("R" & i) = Semi_Finished_Rroduct
                End If
                
            ElseIf Range("D" & i) = "����" Then '�P�_�O����A�h��1
                Range("R" & i) = 1
            End If
            '----------�s�y��----------
            
        Next k
        
        inspector = ""
        Semi_Finished_Rroduct = 0
        
        '----------FQC----------
            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
        
            If Range("D" & i) = "IPQC" Then    '�P�_�O IPQC �B�J�w�Ƥ���""�A�ƻs�@�� IPQC �אּ FQC
                Workbooks(ActWb).Worksheets(1).Activate
                
                For k = 1 To Workbooks(ActWb).Worksheets(1).Cells(1, Columns.Count).End(xlToLeft).Column '�̫�@��
                    If InStr(1, Cells(1, k), "�J�w��") <> 0 Then   '�P�_�O�_����J�w��
                        
                        If Cells(j, k) <> "" Then
                            Semi_Finished_Rroduct = Cells(j, k)
                            
                            Workbooks("�~�OIPQC_FQC����t��(�ե�20210305.xlsm").Worksheets("Q�~���������`��(�[�u)").Activate
                            Range("D" & i & ":AD" & i).Select
                            
                            MsgBox "select = " & Selection.Address
                            
                            Selection.Copy
                            Range("D" & i & ":AD" & i).Offset(1, 0).Select
                            
                            MsgBox "select = " & Selection.Address
                            
                            Selection.PasteSpecial xlPasteValues
                            
                            Range("D" & i).Offset(1, 0) = "FQC"
                            Range("R" & i).Offset(1, 0) = Semi_Finished_Rroduct
        
                        End If
                    End If
                Next k
            End If
            
            Semi_Finished_Rroduct = 0
            '----------FQC----------
        
    Next j
            
End Sub

