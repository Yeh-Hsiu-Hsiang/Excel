Attribute VB_Name = "LTA�w�s�ഫ"
Sub LTA�w�s�ഫ()

    Dim ActWb As String

    ActWb = ActiveWorkbook.Name
    
    '------------------Ū���ö}�ҳ̷s�C��w�s------------------
    Dim pth As String, fn As String, ary(), tmpMax As Long, n As Integer, wb As Workbook

    pth = "\\yeawen\files-server\06_���\01_�ͺ�\��q�C���T\�C��w�s\"    '�]�m���|
    'pth = "C:\Users\ywqa011\Desktop\�C��w�s\"    '�]�m���|
    
    fn = Dir(pth & "*.xls")     '�s����Ƨ��U�� .xls���
    n = 0: tmpMax = 0
    
    Do While fn <> ""
        If fn <> ThisWorkbook.Name Then
            n = n + 1
            ReDim Preserve ary(n)
            ary(n) = Left(Right(fn, 11), 7)   '��J excel �ɦW�����
            If ary(n) > tmpMax Then
                tmpMax = ary(n)   '�̷s����ɮ�
            End If
        End If
        fn = Dir
    Loop
    
    Set wb = Workbooks.Open(pth & "MERP�C��w�s" & tmpMax & ".xls", , True)
    '------------------Ū���ö}�ҳ̷s�C��w�s------------------
    
    
    Workbooks(ActWb).Worksheets("LTA").Activate

    Dim i, j, k As Integer, Find_Value As Long
    
    For k = 8 To Workbooks(ActWb).Worksheets("LTA").Cells(2, Columns.Count).End(xlToLeft).Column '�̫�@��
        If InStr(1, Cells(2, k), Format(Date, "MM/DD")) = 1 Then    '�P�_�O�_���󤵤�
            For i = 3 To Workbooks(ActWb).Worksheets("LTA").Range("C65536").End(xlUp).Row - 1
                For j = 5 To wb.Worksheets("���~�s�q").Range("A65536").End(xlUp).Row
        
                    If Left(wb.Worksheets("���~�s�q").Range("A" & j), 12) = Workbooks(ActWb).Worksheets("LTA").Range("C" & i) Then  '�P�_ LTA �Ƹ����P�󥻤��q�Ƹ�
                        Find_Value = Find_Value + wb.Worksheets("���~�s�q").Range("A" & j).Offset(0, 2)   '�Ƹ��w�s�ƥ[�`
                    End If
                Next j
                
                Workbooks(ActWb).Worksheets("LTA").Activate
                Workbooks(ActWb).Worksheets("LTA").Cells(i, k).Value = Find_Value
                
                Find_Value = 0
            Next i
        End If
    Next k
     
    '------------------���󦡮榡�]�w------------------
    Range("A3:AE9").Select
    Range(Selection, Selection.End(xlDown)).FormatConditions.Delete '�M���榡
    Range(Selection, Selection.End(xlDown)).Select  '����d��
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=OR(ISNUMBER(SEARCH(""-"", $AH9, 1)), ISNUMBER(SEARCH(""-"", $AI9, 1)), ISNUMBER(SEARCH(""-"", $AJ9, 1)))" '�]�w���󤽦�
    
    With Selection.FormatConditions(1).Interior '�]�w�榡
        .PatternColorIndex = xlAutomatic
        .Color = 10066431
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    '------------------���󦡮榡�]�w------------------
    
End Sub

