Sub �פJ�S�w���()

    Dim copyfromfilename, mypath, myfile, endcolumnchar, rang As String
    Dim openfile As Workbook
    Dim endrow, endcolumn, i, j, k As Integer
    
    Application.ScreenUpdating = False
    
    sh1 = Sheets("�Ұʨ��X").Range("b3")
    cd1 = Sheets("�Ұʨ��X").Range("b2")
    
    SK1 = Sheets("�Ұʨ��X").Range("b4")
    

    Application.DisplayAlerts = False
    
    For j = Sheets.Count To 8 Step -1
        Sheets(j).Delete
    Next
    
    Application.DisplayAlerts = True
    
    For k = 7 To 10
        Sheets.Add after:=Sheets(7)  
        ActiveSheet.Name = "no_" & k
    Next k

    copyfromfilename = sh1    '�o�Ӧa��]�w�Q�ƨexcel�ɮ�
    
    mypath = cd1 & "/" '���ɮ׸��|�w�q���ܼ�
    
    myfile = Dir(mypath & "*.xls")   '�̦���M���w���|����*.xls�ɮ�
    
    Do While myfile <> ""

        If myfile = copyfromfilename Then   '���p�M����ݭn�ƻs���ɮ�
        
            Set openfile = Workbooks.Open(mypath & myfile) '�}�ҲŦX�n�D���ɮ�
            
            For i = 1 To openfile.Sheets.Count '�ƻs�Ҧ���sheet
            
                hh = openfile.Sheets.Count
                
                endrow = openfile.Sheets(i).Range("a65536").End(xlUp).Row   '�ھڲĤ@�C�ӽT�w����ƪ��̫�@��
                
                endcolumn = openfile.Sheets(i).Cells(1255).End(xlToLeft).Column '�ھڲĤ@��ӽT�w����ƪ��̫�@�C
                
                endcolumnchar = VBA.Split(Columns(endcolumn).Address, "$")(2)   '���o�̫�@�C�������r��
                
                rang = "a1:AH300" '& endcolumnchar & endrow   'rang = "a1:" & endcolumnchar & endrow   '�c�ئ��зǪ��d��榡 �ҡG��a1�Gc1��'�� rang = "a1:ad300"
                
                openfile.Sheets(i).Range(rang).Copy ThisWorkbook.Sheets(i + 7).Range(rang)   'openfile.Sheets(i).Range(rang).Copy ThisWorkbook.Sheets(i + 7).Range(rang)
                
                
                '----------------
                
                ThisWorkbook.Sheets(i + 7).Name = SK1 & "_" & i   'Sheets(i + 1).Name  .PasteSpecial xlPasteFormats
                
            Next

            Workbooks(myfile).Close False         '�������u�@ï,�ä��@�ק�

        End If

        myfile = Dir

    Loop
    
    '------
    Application.DisplayAlerts = False
    
    For Each ws In Worksheets
        
        If ws.Name Like "no_*" Then    '�P�_�u�@��O�_��No
        
            ws.Delete
        
        End If
    Next
    
    
    Application.DisplayAlerts = True
    '-------
    

    '----------
    Sheets("�Ұʨ��X").Select
    Range("BA2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Formula2R1C1 = _
        "=IF(INDIRECT(RC50&""R""&INT(INT(COLUMN(RC[-17])/36)*COLUMN(RC36)/36)&""C""&IF(MOD(COLUMN(RC[-52]),36)=0,36,MOD(COLUMN(RC[-52]),36)),FALSE)="""","""",INDIRECT(RC50&""R""&INT(INT(COLUMN(RC[-17])/36)*COLUMN(RC36)/36)&""C""&IF(MOD(COLUMN(RC[-52]),36)=0,36,MOD(COLUMN(RC[-52]),36)),FALSE))"
    Range("Q2").Select

End Sub
