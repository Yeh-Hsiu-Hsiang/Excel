Sub 匯入特定資料()

    Dim copyfromfilename, mypath, myfile, endcolumnchar, rang As String
    Dim openfile As Workbook
    Dim endrow, endcolumn, i, j, k As Integer
    
    Application.ScreenUpdating = False
    
    sh1 = Sheets("啟動取出").Range("b3")
    cd1 = Sheets("啟動取出").Range("b2")
    
    SK1 = Sheets("啟動取出").Range("b4")
    

    Application.DisplayAlerts = False
    
    For j = Sheets.Count To 8 Step -1
        Sheets(j).Delete
    Next
    
    Application.DisplayAlerts = True
    
    For k = 7 To 10
        Sheets.Add after:=Sheets(7)  
        ActiveSheet.Name = "no_" & k
    Next k

    copyfromfilename = sh1    '這個地方設定被複制的excel檔案
    
    mypath = cd1 & "/" '把檔案路徑定義給變數
    
    myfile = Dir(mypath & "*.xls")   '依次找尋指定路徑中的*.xls檔案
    
    Do While myfile <> ""

        If myfile = copyfromfilename Then   '假如遍歷到需要複製的檔案
        
            Set openfile = Workbooks.Open(mypath & myfile) '開啟符合要求的檔案
            
            For i = 1 To openfile.Sheets.Count '複製所有的sheet
            
                hh = openfile.Sheets.Count
                
                endrow = openfile.Sheets(i).Range("a65536").End(xlUp).Row   '根據第一列來確定有資料的最後一行
                
                endcolumn = openfile.Sheets(i).Cells(1255).End(xlToLeft).Column '根據第一行來確定有資料的最後一列
                
                endcolumnchar = VBA.Split(Columns(endcolumn).Address, "$")(2)   '取得最後一列對應的字母
                
                rang = "a1:AH300" '& endcolumnchar & endrow   'rang = "a1:" & endcolumnchar & endrow   '構建成標準的範圍格式 例：“a1：c1”'原 rang = "a1:ad300"
                
                openfile.Sheets(i).Range(rang).Copy ThisWorkbook.Sheets(i + 7).Range(rang)   'openfile.Sheets(i).Range(rang).Copy ThisWorkbook.Sheets(i + 7).Range(rang)
                
                
                '----------------
                
                ThisWorkbook.Sheets(i + 7).Name = SK1 & "_" & i   'Sheets(i + 1).Name  .PasteSpecial xlPasteFormats
                
            Next

            Workbooks(myfile).Close False         '關閉源工作簿,並不作修改

        End If

        myfile = Dir

    Loop
    
    '------
    Application.DisplayAlerts = False
    
    For Each ws In Worksheets
        
        If ws.Name Like "no_*" Then    '判斷工作表是否為No
        
            ws.Delete
        
        End If
    Next
    
    
    Application.DisplayAlerts = True
    '-------
    

    '----------
    Sheets("啟動取出").Select
    Range("BA2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Formula2R1C1 = _
        "=IF(INDIRECT(RC50&""R""&INT(INT(COLUMN(RC[-17])/36)*COLUMN(RC36)/36)&""C""&IF(MOD(COLUMN(RC[-52]),36)=0,36,MOD(COLUMN(RC[-52]),36)),FALSE)="""","""",INDIRECT(RC50&""R""&INT(INT(COLUMN(RC[-17])/36)*COLUMN(RC36)/36)&""C""&IF(MOD(COLUMN(RC[-52]),36)=0,36,MOD(COLUMN(RC[-52]),36)),FALSE))"
    Range("Q2").Select

End Sub
