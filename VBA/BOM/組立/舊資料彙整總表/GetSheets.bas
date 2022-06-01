Attribute VB_Name = "GetSheets"
'合併多個Excel檔案
Sub GetSheets()

    Dim Path, ActWb As String, i As Integer
    
    ActWb = ActiveWorkbook.Name         '紀錄現有活頁簿名稱
    
    For i = 6 To Workbooks(ActWb).Worksheets("test").Range("A65536").End(xlUp).Row      '從第六列開始逐筆讀取路徑
        
        Workbooks(ActWb).Worksheets("test").Activate    '指定目前工作表
    
        Path = Range("A" & i)           '路徑
        
        Filename = Dir(Path & "*.xls")      '檔名
        
        Do While Filename <> ""
            Workbooks.Open Filename:=Path & Filename, ReadOnly:=True        '開啟唯讀檔案
        
            '只複製第一個Sheet
            If ActiveWorkbook.Sheets.Count > 0 Then
              ActiveWorkbook.Sheets(1).Copy _
                  after:=ThisWorkbook.Sheets(1)
               ActiveSheet.Name = Filename
            End If
            
            
            Workbooks(ActWb).Worksheets("test").Select
            
            '----------------調整日期格式----------------
            Range("AA:AA, AG:AG, AM:AM").Select
            Selection.NumberFormatLocal = "yyyy/mm/dd"
            '----------------調整日期格式----------------
            
            
            '--------讓選定的位置為有資料的最底行--------
            j = 6
            Do While True
                If ActiveSheet.Cells(j, "C").Value = "" Then
                    ActiveSheet.Cells(j, "C").Select
                    Exit Do
                End If
                j = j + 1
            Loop
            '--------讓選定的位置為有資料的最底行--------
            
            Dim lrow, Version_Row As Long
            
            Worksheets(Filename).Select         '選取匯入的工作表
            
            For n = 10 To Range("A65536").End(xlUp).Row     '從第十列開始讀取直到版本前一列
                
                If Range("A" & n) = "版本" Then
                    lrow = n - 1        '次序列
                    Version_Row = n + 1     '版本列
                End If
            Next
            
            
            '----------------匯入客戶----------------
            Workbooks(ActWb).Worksheets(Filename).Range("D6").Copy
            Workbooks(ActWb).Worksheets("test").Range("B" & j).PasteSpecial xlPasteValues
            '----------------匯入客戶----------------
            
            
            '----------------匯入機種----------------
            Workbooks(ActWb).Worksheets(Filename).Range("G6").Copy
            Workbooks(ActWb).Worksheets("test").Range("C" & j).PasteSpecial xlPasteValues
            '----------------匯入機種----------------
            
            
            Dim Part_No, Lever, Lever1, Lever2, Lever3, Product_Name, Specification, Manufacturer, _
            Dosage, Standard_Loss, ProcessingItem1, ProcessingItem2, ProcessingItem3, _
            ProcessingItem4, ProcessingItem5, ProcessingItem6, ProcessingItem7, _
            ProcessingItem8, ProcessingItem9, Single_Weight, Period, Remark As String
            
            
            For k = 10 To lrow      '匯入所有次序資料
            
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
                
                
                '-----------------Specification 規格-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("I" & k) <> "" Then
                    Specification = Specification & Range("I" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("I" & k) & vbLf
                End If
                '-----------------Specification 規格-----------------
                
                
                '-----------------Manufacturer 廠商-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("J" & k) <> "" Then
                    Manufacturer = Manufacturer & Range("J" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("J" & k) & vbLf
                End If
                '-----------------Manufacturer 廠商-----------------
                
                
                '-----------------Dosage 用量-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("K" & k) <> "" Then
                    Dosage = Dosage & Range("K" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("K" & k) & vbLf
                End If
                '-----------------Dosage 用量-----------------
                
                
                '-----------------Standard_Loss 標準損耗-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("L" & k) <> "" Then
                    Standard_Loss = Standard_Loss & Range("L" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("L" & k) & vbLf
                End If
                '-----------------Standard_Loss 標準損耗-----------------
                
                
                '-----------------ProcessingItem1 加工項目1-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("M" & k) <> "" Then
                    ProcessingItem1 = ProcessingItem1 & Range("M" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("M9") & vbLf
                End If
                '-----------------ProcessingItem1 加工項目1-----------------
                
                
                '-----------------ProcessingItem2 加工項目2-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("N" & k) <> "" Then
                    ProcessingItem2 = ProcessingItem2 & Range("N" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("N9") & vbLf
                End If
                '-----------------ProcessingItem2 加工項目2-----------------
                
                
                '-----------------ProcessingItem3 加工項目3-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("O" & k) <> "" Then
                    ProcessingItem3 = ProcessingItem3 & Range("O" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("O9") & vbLf
                End If
                '-----------------ProcessingItem3 加工項目3-----------------
                
                
                '-----------------ProcessingItem4 加工項目4-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("P" & k) <> "" Then
                    ProcessingItem4 = ProcessingItem4 & Range("P" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("P9") & vbLf
                End If
                '-----------------ProcessingItem4 加工項目4-----------------
                
                
                '-----------------ProcessingItem5 加工項目5-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("Q" & k) <> "" Then
                    ProcessingItem5 = ProcessingItem5 & Range("Q" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("Q9") & vbLf
                End If
                '-----------------ProcessingItem5 加工項目5-----------------
                
                
                '-----------------ProcessingItem6 加工項目6-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("R" & k) <> "" Then
                    ProcessingItem6 = ProcessingItem6 & Range("R" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("R9") & vbLf
                End If
                '-----------------ProcessingItem6 加工項目6-----------------
                
                
                '-----------------ProcessingItem7 加工項目7-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("S" & k) <> "" Then
                    ProcessingItem7 = ProcessingItem7 & Range("S" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("S9") & vbLf
                End If
                '-----------------ProcessingItem7 加工項目7-----------------
                
                
                '-----------------ProcessingItem8 加工項目8-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("T" & k) <> "" Then
                    ProcessingItem8 = ProcessingItem8 & Range("T" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("T9") & vbLf
                End If
                '-----------------ProcessingItem8 加工項目8-----------------
                
                
                '-----------------ProcessingItem9 加工項目9-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("U" & k) <> "" Then
                    ProcessingItem9 = ProcessingItem9 & Range("U" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("U9") & vbLf
                End If
                '-----------------ProcessingItem9 加工項目9-----------------
                
                
                '-----------------Single_Weight 單重-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("V" & k) <> "" Then
                    Single_Weight = Single_Weight & Range("V" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("V" & k) & vbLf
                End If
                '-----------------Single_Weight 單重-----------------
                
                
                '-----------------Period 週期-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("X" & k) <> "" Then
                    Period = Period & Range("X" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("X" & k) & vbLf
                End If
                '-----------------Period 週期-----------------
                
                
                '-----------------Remark 備註-----------------
                If Workbooks(ActWb).Worksheets(Filename).Range("Y" & k) <> "" Then
                    Remark = Remark & Range("Y" & k).Address & "_" & Workbooks(ActWb).Worksheets(Filename).Range("Y" & k) & vbLf
                End If
                '-----------------Remark 備註-----------------
            Next
            
            
            Worksheets(Filename).Select
            
        
            '-----------------Version1 版本1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("Z" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Version1 版本1-----------------
            
            
            
            '-----------------date1 製作 / 修訂日期1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AA" & j).PasteSpecial xlPasteValues
            End If
            '-----------------date1 製作 / 修訂日期1-----------------
            
            
            
            '-----------------ChangeRecord1 備註 / 變更記錄1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AB" & j).PasteSpecial xlPasteValues
            End If
            '-----------------ChangeRecord1 備註 / 變更記錄1-----------------
            
            
            
            '-----------------Approve1 核準1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AC" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Approve1 核準1-----------------
            
            
            '-----------------Review1 審核1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AD" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Review1 審核1-----------------
            
            
            '-----------------Tabulation1 製表1-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row).Copy
                Workbooks(ActWb).Worksheets("test").Range("AE" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Tabulation1 製表1-----------------
            
            
            
            '-----------------Version2 版本2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AF" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Version2 版本2-----------------
            
            
            
            '-----------------date2 製作 / 修訂日期2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AG" & j).PasteSpecial xlPasteValues
            End If
            '-----------------date2 製作 / 修訂日期2-----------------
            
            
            
            '-----------------ChangeRecord2 備註 / 變更記錄2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AH" & j).PasteSpecial xlPasteValues
            End If
            '-----------------ChangeRecord2 備註 / 變更記錄2-----------------
            
            
            
            '-----------------Approve2 核準2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AI" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Approve2 核準2-----------------
            
            
            '-----------------Review2 審核2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AJ" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Review2 審核2-----------------
            
            
            '-----------------Tabulation2 製表2-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row + 1) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row + 1).Copy
                Workbooks(ActWb).Worksheets("test").Range("AK" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Tabulation2 製表2-----------------
            

            '-----------------Version3 版本3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("A" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AL" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Version3 版本3-----------------
            
            
            
            '-----------------date3 製作 / 修訂日期3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("C" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AM" & j).PasteSpecial xlPasteValues
            End If
            '-----------------date3 製作 / 修訂日期3-----------------
            
            
            
            '-----------------ChangeRecord3 備註 / 變更記錄3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("F" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AN" & j).PasteSpecial xlPasteValues
            End If
            '-----------------ChangeRecord3 備註 / 變更記錄3-----------------
            
            
            
            '-----------------Approve3 核準3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("M" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AO" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Approve3 核準3-----------------
            
            
            '-----------------Review3 審核3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("Q" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AP" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Review3 審核3-----------------
            
            
            '-----------------Tabulation3 製表3-----------------
            If Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row + 2) <> "" Then
                Workbooks(ActWb).Worksheets(Filename).Range("W" & Version_Row + 2).Copy
                Workbooks(ActWb).Worksheets("test").Range("AQ" & j).PasteSpecial xlPasteValues
            End If
            '-----------------Tabulation3 製表3-----------------
            
            
            
            '-------------------存入test工作表-------------------
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

            '-------------------清空暫存-------------------
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
            '-------------------清空暫存-------------------
            
            Application.CutCopyMode = False     '取消選取複製欄位
            
            Application.DisplayAlerts = False
            Worksheets(Filename).Delete
            Application.DisplayAlerts = True
        
            Workbooks(Filename).Close       '關閉唯讀檔案
            Filename = Dir()
          Loop
    Next i
    
End Sub


