Attribute VB_Name = "GetSheetsFromOther"
'從路徑清單中開啟並複製 Excel 工作表到一個檔案
Sub GetSheetsFromOther()

    Dim Path, ActWb As String, i As Integer
    
    ActWb = ActiveWorkbook.Name         '紀錄現有活頁簿名稱
    
    For i = 2 To Workbooks(ActWb).Worksheets("目錄").Range("A65536").End(xlUp).Row      '逐筆讀取路徑
        
        Workbooks(ActWb).Worksheets("目錄").Activate    '指定目前工作表
    
        Path = Range("A" & i)           '路徑
        
        fileName = Dir(Path & "*.xls")      '檔名
        
        Do While fileName <> ""
            Workbooks.Open fileName:=Path & fileName, ReadOnly:=True        '開啟唯讀檔案
            OpenWb = ActiveWorkbook.Name
            
            For x = 1 To Workbooks(OpenWb).Sheets.Count
                If Workbooks(OpenWb).Sheets.Count > 1 Then
                    If Range("Q2") = "" Then
                        ActiveWorkbook.Sheets(x).Copy _
                        After:=Workbooks(ActWb).Sheets(1)
                        MsgBox x
                        ActiveSheet.Name = Range("K4") & "#" & Range("O5") & "-" & x  '工作表名稱
                    Else
                        GoTo ContinueForLoop
                    End If
                                
                Else
                    ActiveWorkbook.Sheets(x).Copy _
                    After:=Workbooks(ActWb).Sheets(1)
                    ActiveSheet.Name = Range("K4") & "#" & Range("O5")  '工作表名稱
                End If
ContinueForLoop:
            Next x

            Workbooks(fileName).Close       '關閉唯讀檔案
            fileName = Dir()
          Loop
    Next i
    
    
    Sheets("目錄").Select
    
    Columns_B
    
    Columns_C_D
    
    New_Sheets
    
    GetNewSheets
    
    Del_list
    
End Sub

Sub Columns_B()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        B_Position = 1
        Do While True
            If ActiveSheet.Cells(B_Position, "B").Value = "" Then
                ActiveSheet.Cells(B_Position, "B").Select
                Exit Do
            End If
            B_Position = B_Position + 1
        Loop
    
        Cells(B_Position, "B") = ws.Name
    Next
    
End Sub

Sub Columns_C_D()

    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    
    Range("C2").FormulaR1C1 = "=IFERROR(LEFT(RC[-1], FIND(""#"", RC[-1])-1),"""")"
    Range("C2").AutoFill Destination:=Range("C2:C" & lrow)
    Range("C2:C" & lrow).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues

    Range("D2").FormulaR1C1 = "=IFERROR(MID(RC[-2], FIND(""#"", RC[-2])+1, LEN(RC[-2])-FIND(""#"", RC[-2])+1),"""")"
    Range("D2").AutoFill Destination:=Range("D2:D" & lrow)
    Range("D2:D" & lrow).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    
    Columns("C:D").Select
    Columns("C:D").EntireColumn.AutoFit
    
    Application.CutCopyMode = False

End Sub

Sub New_Sheets()

    Dim ws As Worksheet

    For C_Loop = 2 To Range("C65536").End(xlUp).Row
    
        If Range("C" & C_Loop) = Range("C" & C_Loop).Offset(-1, 0) And Range("C" & C_Loop) <> "" Then
            If Range("D" & C_Loop) < Range("D" & C_Loop).Offset(-1, 0) And Range("D" & C_Loop) <> "" Then

                For Each ws In Worksheets
                
                    If LCase(ws.Name) = LCase(Range("B" & C_Loop)) Then   '判斷是否已存在工作表
                    
                        Application.DisplayAlerts = False
                        Sheets(LCase(Range("B" & C_Loop))).Delete
                        Application.DisplayAlerts = True
                        
                        Range("C" & C_Loop, "D" & C_Loop).Select
                        Selection.Clear
                        
                        C_Loop = C_Loop - 1
                    End If
                Next
            End If
        End If
    Next


End Sub
Sub GetNewSheets()

    Dim ws As Worksheet
    
    Cells(1, "F") = "目錄"
    i = 1
    
    For Each ws In Worksheets
        If ws.Name <> "目錄" Then
            i = i + 1
            
            ActiveSheet.Hyperlinks.Add anchor:=Cells(i, "F"), _
                                       Address:="", _
                                       SubAddress:="'" & ws.Name & "'!A1", _
                                       TextToDisplay:=Worksheets(ws.Name).Range("K4") & "產品檢驗規範(加工檢驗)-(" & Worksheets(ws.Name).Range("M2") & ")"

'            ws.Hyperlinks.Add anchor:=ws.Cells(1, "P"), _
'                              Address:="", _
'                              SubAddress:="目錄!A1", _
'                              TextToDisplay:="返回目錄"
'                              ws.Cells(1, "P").Font.Size = 16
'                              ws.Cells(1, "P").EntireColumn.AutoFit
        End If
    Next
    
    With Worksheets("目錄").Range("F:F")
        .Font.Size = 20
    End With

    Columns("F:F").EntireColumn.AutoFit

End Sub

Sub Del_list()

    SendKeys "^g^a{DEL}"
    
    Range("A2", Range("A65535").End(xlUp)).Select
    Selection.ClearContents
End Sub


