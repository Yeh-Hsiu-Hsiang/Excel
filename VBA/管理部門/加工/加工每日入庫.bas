Attribute VB_Name = "�[�u�C��J�w"
Sub �[�u�C��J�w()

    ActWb = ActiveWorkbook.Name
    
    '-----------�s�O���-----------
    Range("L2").Select
    ActiveCell.Formula = "=TEXT(B2, ""m��d��"")"
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

    Selection.AutoFill Destination:=Range("L2:L" & lrow)
    Range("L2:L" & lrow).Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    '-----------�s�O���-----------
    
    ActiveSheet.Range("A:K").AutoFilter Field:=7, Criteria1:="1"    '�z���渹��1�����
    Cells.Replace What:=".0000", Replacement:="", LookAt:=xlPart    '�p���I��|����N
    Range("E:F, H:K").Select    '�⤣�ݭn���������
    Selection.EntireColumn.Hidden = True
    
    '-----------��������m-----------
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("D:D").Select
    Selection.Cut
    Range("A1").Select
    ActiveSheet.Paste
    Columns("E:E").Select
    Selection.Cut
    Range("D1").Select
    ActiveSheet.Paste
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    '-----------��������m-----------
    
    Range("A2", Range("D65536").End(xlUp)).Select
    Selection.Copy

    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\�ѫ�\0617_�[�u�C��J�w.xls"  '�}���ɮ�
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Name Like "*�~" Then
            ws.Activate
            Exit For
        End If
    Next ws

    i = 3
    Do While True
        If ActiveSheet.Cells(i, 1).Value = "" Then
            ActiveSheet.Cells(i, 1).Select
            Exit Do
        End If
        i = i + 1
    Loop
    
    ActiveSheet.Select
    Selection.PasteSpecial xlPasteValues
    
End Sub
