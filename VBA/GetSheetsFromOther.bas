Attribute VB_Name = "GetSheetsFromOther"
'�q���|�M�椤�}�Ҩýƻs Excel �u�@���@���ɮ�
Sub GetSheetsFromOther()

    Dim Path, ActWb As String, i As Integer
    
    ActWb = ActiveWorkbook.Name         '�����{������ï�W��
    
    For i = 2 To Workbooks(ActWb).Worksheets("test").Range("A65536").End(xlUp).Row      '�v��Ū�����|
        
        Workbooks(ActWb).Worksheets("test").Activate    '���w�ثe�u�@��
    
        Path = Range("A" & i)           '���|
        
        fileName = Dir(Path & "*.xls")      '�ɦW
        
        Do While fileName <> ""
            Workbooks.Open fileName:=Path & fileName, ReadOnly:=True        '�}�Ұ�Ū�ɮ�
            OpenWb = ActiveWorkbook.Name
            
            For x = 1 To Workbooks(OpenWb).Sheets.Count
                If Workbooks(OpenWb).Sheets.Count > 1 Then
                    ActiveWorkbook.Sheets(x).Copy _
                    After:=Workbooks(ActWb).Sheets(1)
                    ActiveSheet.Name = Range("K4") & "#" & Range("O5") & "-" & x  '�u�@��W��
                Else
                    ActiveWorkbook.Sheets(x).Copy _
                    After:=Workbooks(ActWb).Sheets(1)
                    ActiveSheet.Name = Range("K4") & "#" & Range("O5")  '�u�@��W��
                End If
            Next

            Workbooks(fileName).Close       '������Ū�ɮ�
            fileName = Dir()
          Loop
    Next i
    
    Sheets("test").Select
    
    Columns_B
    
    Columns_C_D
    
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

    For C_Loop = 2 To Range("C65536").End(xlUp).Row
    
        If Range("C" & C_Loop) = Range("C" & C_Loop).Offset(-1, 0) And Range("C" & C_Loop) <> "" Then
            If Range("D" & C_Loop) < Range("D" & C_Loop).Offset(-1, 0) And Range("D" & C_Loop) <> "" Then
                'Rows(C_Loop).Select
                'Selection.Delete Shift:=xlUp
                
                Range("C" & C_Loop, "D" & C_Loop).Select
                Selection.Clear
                
                C_Loop = C_Loop - 1
            
            End If
  
        End If
    Next


End Sub


Sub Del_list()
       SendKeys "^g^a{DEL}"
End Sub

