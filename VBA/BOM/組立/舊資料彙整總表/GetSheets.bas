Attribute VB_Name = "GetSheets"
'�X�֦h��Excel�ɮ�
Sub GetSheets()

    Dim Path, ActWb As String, i As Integer
    
    ActWb = ActiveWorkbook.Name         '�����{������ï�W��
    
    For i = 6 To Workbooks(ActWb).Worksheets("test").Range("A65536").End(xlUp).Row      '�q�Ĥ��C�}�l�v��Ū�����|
        
        Workbooks(ActWb).Worksheets("test").Activate    '���w�ثe�u�@��
    
        Path = Range("A" & i)           '���|
        
        Filename = Dir(Path & "*.xls")      '�ɦW
        
        Do While Filename <> ""
            Workbooks.Open Filename:=Path & Filename, ReadOnly:=True        '�}�Ұ�Ū�ɮ�
        
            '�u�ƻs�Ĥ@��Sheet
            If ActiveWorkbook.Sheets.Count > 0 Then
              ActiveWorkbook.Sheets(1).Copy _
                  after:=ThisWorkbook.Sheets(1)
               ActiveSheet.Name = Filename
            End If
            
            
            Workbooks(ActWb).Worksheets("test").Select
            
            '----------------�վ����榡----------------
            Range("AA:AA, AG:AG, AM:AM").Select
            Selection.NumberFormatLocal = "yyyy/mm/dd"
            '----------------�վ����榡----------------
            
            
            '--------����w����m������ƪ��̩���--------
            j = 6
            Do While True
                If ActiveSheet.Cells(j, "C").Value = "" Then
                    ActiveSheet.Cells(j, "C").Select
                    Exit Do
                End If
                j = j + 1
            Loop
            '--------����w����m������ƪ��̩���--------
            
            Dim lrow, Version_Row As Long
            
            Worksheets(Filename).Select         '����פJ���u�@��
            
            For n = 10 To Range("A65536").End(xlUp).Row     '�q�ĤQ�C�}�lŪ�����쪩���e�@�C
                
                If Range("A" & n) = "����" Then
                    lrow = n - 1        '���ǦC
                    Version_Row = n + 1     '�����C
                End If
            Next
            
            
            '----------------�פJ�Ȥ�----------------
            Workbooks(ActWb).Worksheets(Filename).Range("D6").Copy
            Workbooks(ActWb).Worksheets("test").Range("B" & j).PasteSpecial xlPasteValues
            '----------------�פJ�Ȥ�----------------
            
            
            '----------------�פJ����----------------
            Workbooks(ActWb).Worksheets(Filename).Range("G6").Copy
            Workbooks(ActWb).Worksheets("test").Range("C" & j).PasteSpecial xlPasteValues
            '----------------�פJ����----------------
            
            
            Dim Part_No, Lever, Lever1, Lever2, Lever3, Product_Name, Specification, Manufacturer, _
            Dosage, Standard_Loss, ProcessingItem1, ProcessingItem2, ProcessingItem3, _
            ProcessingItem4, ProcessingItem5, ProcessingItem6, ProcessingItem7, _
            ProcessingItem8, ProcessingItem9, Single_Weight, Period, Remark As String
            
            
            For k = 10 To lrow      '�פJ�Ҧ����Ǹ��
            
                '-----------------Part_No-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("C" & k) <> "" Then
                    Part_No = Part_No & Range("C" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("C" & k) & vbLf
                End If
                '-----------------Part_No-----------------
            
            
                '-----------------Lever-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("D" & k) <> "" Then
                    Lever = Lever & Range("D" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("D" & k) & vbLf
                End If
                '-----------------Lever-----------------
            
            
                '-----------------Lever1-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("E" & k) <> "" Then
                    Lever1 = Lever1 & Range("E" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("E" & k) & vbLf
                End If
                '-----------------Lever1-----------------
                
                
                '-----------------Lever2-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("F" & k) <> "" Then
                    Lever2 = Lever2 & Range("F" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("F" & k) & vbLf
                End If
                '-----------------Lever2-----------------
                
                
                '-----------------Lever3-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("G" & k) <> "" Then
                    Lever3 = Lever3 & Range("G" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("G" & k) & vbLf
                End If
                '-----------------Lever3-----------------
                
                
                '-----------------Product_Name-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("H" & k) <> "" Then
                    Product_Name = Product_Name & Range("H" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("H" & k) & vbLf
                End If
                '-----------------Product_Name-----------------
                
                
                '-----------------Specification �W��-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("I" & k) <> "" Then
                    Specification = Specification & Range("I" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("I" & k) & vbLf
                End If
                '-----------------Specification �W��-----------------
                
                
                '-----------------Manufacturer �t��-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("J" & k) <> "" Then
                    Manufacturer = Manufacturer & Range("J" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("J" & k) & vbLf
                End If
                '-----------------Manufacturer �t��-----------------
                
                
                '-----------------Dosage �ζq-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("K" & k) <> "" Then
                    Dosage = Dosage & Range("K" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("K" & k) & vbLf
                End If
                '-----------------Dosage �ζq-----------------
                
                
                '-----------------Standard_Loss �зǷl��-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("L" & k) <> "" Then
                    Standard_Loss = Standard_Loss & Range("L" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("L" & k) & vbLf
                End If
                '-----------------Standard_Loss �зǷl��-----------------
                
                
                '-----------------ProcessingItem1 �[�u����1-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("M" & k) <> "" Then
                    ProcessingItem1 = ProcessingItem1 & Range("M" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("M9") & vbLf
                End If
                '-----------------ProcessingItem1 �[�u����1-----------------
                
                
                '-----------------ProcessingItem2 �[�u����2-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("N" & k) <> "" Then
                    ProcessingItem2 = ProcessingItem2 & Range("N" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("N9") & vbLf
                End If
                '-----------------ProcessingItem2 �[�u����2-----------------
                
                
                '-----------------ProcessingItem3 �[�u����3-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("O" & k) <> "" Then
                    ProcessingItem3 = ProcessingItem3 & Range("O" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("O9") & vbLf
                End If
                '-----------------ProcessingItem3 �[�u����3-----------------
                
                
                '-----------------ProcessingItem4 �[�u����4-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("P" & k) <> "" Then
                    ProcessingItem4 = ProcessingItem4 & Range("P" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("P9") & vbLf
                End If
                '-----------------ProcessingItem4 �[�u����4-----------------
                
                
                '-----------------ProcessingItem5 �[�u����5-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("Q" & k) <> "" Then
                    ProcessingItem5 = ProcessingItem5 & Range("Q" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("Q9") & vbLf
                End If
                '-----------------ProcessingItem5 �[�u����5-----------------
                
                
                '-----------------ProcessingItem6 �[�u����6-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("R" & k) <> "" Then
                    ProcessingItem6 = ProcessingItem6 & Range("R" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("R9") & vbLf
                End If
                '-----------------ProcessingItem6 �[�u����6-----------------
                
                
                '-----------------ProcessingItem7 �[�u����7-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("S" & k) <> "" Then
                    ProcessingItem7 = ProcessingItem7 & Range("S" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("S9") & vbLf
                End If
                '-----------------ProcessingItem7 �[�u����7-----------------
                
                
                '-----------------ProcessingItem8 �[�u����8-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("T" & k) <> "" Then
                    ProcessingItem8 = ProcessingItem8 & Range("T" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("T9") & vbLf
                End If
                '-----------------ProcessingItem8 �[�u����8-----------------
                
                
                '-----------------ProcessingItem9 �[�u����9-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("U" & k) <> "" Then
                    ProcessingItem9 = ProcessingItem9 & Range("U" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("U9") & vbLf
                End If
                '-----------------ProcessingItem9 �[�u����9-----------------
                
                
                '-----------------Single_Weight �歫-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("V" & k) <> "" Then
                    Single_Weight = Single_Weight & Range("V" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("V" & k) & vbLf
                End If
                '-----------------Single_Weight �歫-----------------
                
                
                '-----------------Period �g��-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("X" & k) <> "" Then
                    Period = Period & Range("X" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("X" & k) & vbLf
                End If
                '-----------------Period �g��-----------------
                
                
                '-----------------Remark �Ƶ�-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("Y" & k) <> "" Then
                    Remark = Remark & Range("Y" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("Y" & k) & vbLf
                End If
                '-----------------Remark �Ƶ�-----------------
            Next
            
            
            Worksheets(Filename).Select
            
        
            '-----------------Version1 ����1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("Z" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Version1 ����1-----------------
            
            
            
            '-----------------date1 �s�@ / �׭q���1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AA" & j).PasteSpecial xlPasteValues
            End If
            '-----------------date1 �s�@ / �׭q���1-----------------
            
            
            
            '-----------------ChangeRecord1 �Ƶ� / �ܧ�O��1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AB" & j).PasteSpecial xlPasteValues
            End If
            '-----------------ChangeRecord1 �Ƶ� / �ܧ�O��1-----------------
            
            
            
            '-----------------Approve1 �ַ�1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AC" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Approve1 �ַ�1-----------------
            
            
            '-----------------Review1 �f��1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AD" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Review1 �f��1-----------------
            
            
            '-----------------Tabulation1 �s��1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AE" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Tabulation1 �s��1-----------------
            
            
            
            '-----------------Version2 ����2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AF" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Version2 ����2-----------------
            
            
            
            '-----------------date2 �s�@ / �׭q���2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AG" & j).PasteSpecial xlPasteValues
            End If
            '-----------------date2 �s�@ / �׭q���2-----------------
            
            
            
            '-----------------ChangeRecord2 �Ƶ� / �ܧ�O��2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AH" & j).PasteSpecial xlPasteValues
            End If
            '-----------------ChangeRecord2 �Ƶ� / �ܧ�O��2-----------------
            
            
            
            '-----------------Approve2 �ַ�2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AI" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Approve2 �ַ�2-----------------
            
            
            '-----------------Review2 �f��2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AJ" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Review2 �f��2-----------------
            
            
            '-----------------Tabulation2 �s��2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AK" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Tabulation2 �s��2-----------------
            

            '-----------------Version3 ����3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AL" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Version3 ����3-----------------
            
            
            
            '-----------------date3 �s�@ / �׭q���3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AM" & j).PasteSpecial xlPasteValues
            End If
            '-----------------date3 �s�@ / �׭q���3-----------------
            
            
            
            '-----------------ChangeRecord3 �Ƶ� / �ܧ�O��3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AN" & j).PasteSpecial xlPasteValues
            End If
            '-----------------ChangeRecord3 �Ƶ� / �ܧ�O��3-----------------
            
            
            
            '-----------------Approve3 �ַ�3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AO" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Approve3 �ַ�3-----------------
            
            
            '-----------------Review3 �f��3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AP" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Review3 �f��3-----------------
            
            
            '-----------------Tabulation3 �s��3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AQ" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Tabulation3 �s��3-----------------
            
            
            
            '-------------------�s�Jtest�u�@��-------------------
            Workbooks(ActWb).Worksheets("test").Range("D" & j) = Part_No
            Workbooks(ActWb).Worksheets("test").Range("E" & j) = Lever
            Workbooks(ActWb).Worksheets("test").Range("F" & j) = Lever1
            Workbooks(ActWb).Worksheets("test").Range("G" & j) = Lever2
            Workbooks(ActWb).Worksheets("test").Range("H" & j) = Lever3
            Workbooks(ActWb).Worksheets("test").Range("I" & j) = Product_Name
            Workbooks(ActWb).Worksheets("test").Range("J" & j) = Specification
            Workbooks(ActWb).Worksheets("test").Range("K" & j) = Manufacturer
            Workbooks(ActWb).Worksheets("test").Range("L" & j) = Dosage
            Workbooks(ActWb).Worksheets("test").Range("M" & j) = Standard_Loss
            Workbooks(ActWb).Worksheets("test").Range("N" & j) = ProcessingItem1
            Workbooks(ActWb).Worksheets("test").Range("O" & j) = ProcessingItem2
            Workbooks(ActWb).Worksheets("test").Range("P" & j) = ProcessingItem3
            Workbooks(ActWb).Worksheets("test").Range("Q" & j) = ProcessingItem4
            Workbooks(ActWb).Worksheets("test").Range("R" & j) = ProcessingItem5
            Workbooks(ActWb).Worksheets("test").Range("S" & j) = ProcessingItem6
            Workbooks(ActWb).Worksheets("test").Range("T" & j) = ProcessingItem7
            Workbooks(ActWb).Worksheets("test").Range("U" & j) = ProcessingItem8
            Workbooks(ActWb).Worksheets("test").Range("V" & j) = ProcessingItem9
            Workbooks(ActWb).Worksheets("test").Range("W" & j) = Single_Weight
            Workbooks(ActWb).Worksheets("test").Range("X" & j) = Period
            Workbooks(ActWb).Worksheets("test").Range("Y" & j) = Remark

            '-------------------�M�żȦs-------------------
            Part_No = ""
            Lever = ""
            Lever1 = ""
            Lever2 = ""
            Lever3 = ""
            Product_Name = ""
            Specification = ""
            Manufacturer = ""
            Dosage = ""
            Standard_Loss = ""
            ProcessingItem1 = ""
            ProcessingItem2 = ""
            ProcessingItem3 = ""
            ProcessingItem4 = ""
            ProcessingItem5 = ""
            ProcessingItem6 = ""
            ProcessingItem7 = ""
            ProcessingItem8 = ""
            ProcessingItem9 = ""
            Single_Weight = ""
            Period = ""
            Remark = ""
            '-------------------�M�żȦs-------------------
            
            Application.CutCopyMode = False     '��������ƻs���
            
            Application.DisplayAlerts = False
            Worksheets(Filename).Delete
            Application.DisplayAlerts = True
        
            Workbooks(Filename).Close       '������Ū�ɮ�
            Filename = Dir()
          Loop
    Next i
    
End Sub


