Sub �ץXTEST�P�f()

    Dim wb As String

    wb = ActiveWorkbook.Name
    
    Workbooks.Open Filename:="\\yeawen\files-server\10_����\00_i-Reporter ��ʪ��t��\����\���X�ӤH����\�P�f\TEST�P�f.xls"  '�}���ɮ�
    'Workbooks.Open Filename:="C:\Users\ywqa011\Desktop\���X\�P�f\TEST�P�f.xls"  '�}���ɮ�
    
    Range("A2", "AX" & Range("A65536").End(xlUp).Row + 1).Select
    Selection.Delete
    
    Workbooks(wb).Worksheets("��P�f������").Activate
    
    Range("A2", "AX" & Range("A65536").End(xlUp).Row).Select
    Selection.Copy
    
    Workbooks("TEST�P�f.xls").Worksheets(1).Activate
    Range("A2").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
End Sub
