Attribute VB_Name = "�ק�"
Sub �ק�()
    Application.ScreenUpdating = False

    Set get_E2 = Sheets("��J").Range("H2")
    
    
    If Sheets("��J").Range("E2") = "" Then
        
        Sheets("��J").Range("AI4:HO4").Select
        Application.CutCopyMode = False
        Selection.Copy
        
        Sheets("�Ȥ�D��").Select
        Range("A1").Select
        
        '----------����w����m������ƪ��̩���----------
        i = 3
        Do While True
        If ActiveSheet.Cells(i, "G").Value = "" Then
            ActiveSheet.Cells(i, "G").Select
            Exit Do
        End If
        i = i + 1
        Loop
        '----------����w����m������ƪ��̩���----------
        
        Selection.PasteSpecial Paste:=xlPasteValues
        
        Range("A3").Select
        
        Sheets("��J").Select
        Range("E2").Select
    
    Else
        If Range("E2") <> "" Then
        
            Sheets("��J").Range("AI4:HO4").Select
            Application.CutCopyMode = False
            Selection.Copy
            
            
            Sheets("�Ȥ�D��").Select
            
            ActiveSheet.Cells(get_E2, "A").Select
            
            Selection.PasteSpecial Paste:=xlPasteValues
            
            Application.CutCopyMode = False
            
            Range("A1").Select
            Sheets("��J").Select
            Range("E2").Select
            
        End If
    End If
    
    For k = 35 To Workbooks(ActWb).Worksheets("��J").Cells(5, Columns.Count).End(xlToLeft).Column '�̫�@��
    
        If Cells(5, k) <> "" Then
        
        End If
    
    
    Next
    
    
    �M���Ȥ����
    
End Sub

Sub �M���Ȥ����()

    Sheets("��J").Select
    Range("E2, E5, I5, D8, F8, J8, L8, S8, Y8, E11, H11, K11, P11").Select
    Application.CutCopyMode = False
    Selection.ClearContents

    '--------���h����--------
    For i = 28 To ActiveSheet.Range("C65536").End(xlUp).Row + 1 Step 2
        Range("C" & i & ":Z" & i).Select
        Selection.ClearContents
    Next i
    '--------���h����--------
    
    '--------����--------
    For j = 15 To 23 Step 2
        Range("C" & j & ":Z" & j).Select
        Selection.ClearContents
    Next j
    '--------����--------
    
    
    '------------�R��BOM�B���~�ϡBFA------------
    Range("D117:F117").Select
    Selection.ClearContents
    '------------�R��BOM�B���~�ϡBFA------------
    
    
    '------------�R���s���------------
    For k = 4 To 16
        '------------�R���s���1~10------------
        Cells(121, k).Select
        Selection.ClearContents
        '------------�R���s���1~10------------
        
        
        '------------�R���������1~10------------
        Cells(124, k).Select
        Selection.ClearContents
        '------------�R���������1~10------------


        '------------�R���s���11~20------------
        Cells(128, k).Select
        Selection.ClearContents
        '------------�R���s���11~20------------


        '------------�R���������11~20------------
        Cells(131, k).Select
        Selection.ClearContents
        '------------�R���������11~20------------


        '------------�R���s���21~30------------
        Cells(135, k).Select
        Selection.ClearContents
        '------------�R���s���21~30------------


        '------------�R���������21~30------------
        Cells(138, k).Select
        Selection.ClearContents
        '------------�R���������21~30------------
    Next
    '------------�R���s���------------
    
    
    
    '------------�R�����~------------
    Range("D144").Select
    Selection.ClearContents
    '------------�R�����~------------
    
    
    '------------�R���s��------------
    For l = 4 To 16
        '------------�R���s��1~10------------
        Cells(148, l).Select
        Selection.ClearContents
        '------------�R���s��1~10------------
        
        
        '------------�R���������1~10------------
        Cells(151, k).Select
        Selection.ClearContents
        '------------�R���������1~10------------


        '------------�R���s��11~20------------
        Cells(155, k).Select
        Selection.ClearContents
        '------------�R���s��11~20------------


        '------------�R���������11~20------------
        Cells(158, k).Select
        Selection.ClearContents
        '------------�R���������11~20------------


        '------------�R���s��21~30------------
        Cells(162, k).Select
        Selection.ClearContents
        '------------�R���s��21~30------------


        '------------�R���������21~30------------
        Cells(165, k).Select
        Selection.ClearContents
        '------------�R���������21~30------------
    Next
    '------------�R���s��------------
    
    Range("E2").Select
End Sub


