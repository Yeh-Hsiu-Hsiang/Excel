Attribute VB_Name = "������"

Sub ������()
Attribute ������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������ ����
'
' �ֳt��: Ctrl+w
'
Range("A2:C2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("A7:BH7").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=6
    Selection.Copy
       
    
    Sheets("�������").Select
    
    INN7
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    '--------- ������ƨ��`��AF ---------
    Workbooks.Open Filename:="\\yeawen\files-server\02_����\�����C���T\���������\����_�����Ͳ��`����_AF.xlsm"
    
    Sheets("���������").Select
    
    INN7
    
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("G:R, V:AA, AE:AM, AO:BH").ClearContents   '�M����Ƥ��e
    '--------- ������ƨ��`��AF ---------
        
    
    
    MsgBox "����ഫ����!"
    
End Sub
Sub INN7()

'����w����m������ƪ��̩���

i = 7
    Do While True
        If ActiveSheet.Cells(i, 1).Value = "" Then
            ActiveSheet.Cells(i, 1).Select
            Exit Do
        End If
        i = i + 1
    Loop
  
    
End Sub
Sub ����0�C()

Application.ScreenUpdating = False

ASD = ActiveCell.Column


Rows("78:105").Select

Selection.EntireRow.Hidden = False

For i = 78 To 105

If Cells(i, ASD) = "0" Then

Rows(i).Hidden = True

End If

Next

End Sub

