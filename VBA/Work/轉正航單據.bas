Sub 轉正航單據()
'
' 轉正航單據 巨集
'
Set YEE = Sheets("訂單").Range("S1")

Set GEE = Sheets("OSP轉正航單據").Range("B1")

Set BEE = Sheets("OSP").Range("R1")

'
MsgBox "  *** 現在要將訂單明細 - 轉成[ 正航 ] 訂單憑證單據 ***  "
    Sheets("RD訂單單據轉出").Select
    Range("A2:AY600").Select
    Selection.ClearContents

    Range("A2").Select
    Sheets("訂單轉正航單據").Select
    Range("A2:AY2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("RD訂單單據轉出").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    '-------轉OSP
    Sheets("OSP轉正航單據").Select
    Range("C2:BA" & BEE).Select
   ' Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("RD訂單單據轉出").Select
    Range("A" & GEE).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
   '----------列印單價差異明細
   Sheets("單價").Select
   If Range("U1") > 1 Or Range("AF1") > 1 Then
        
      MsgBox "  ****即將列印單價差異明細表****  "
      
      
    印單價差異
   End If
   '---------
   MsgBox "  @@@ 即將實施 轉正航單據檔案  @@@  "

    Sheets("RD訂單單據轉出").Select

    For i = 2 To Range("A65536").End(xlUp).Row
        If Range("A" & i) = Range("A" & i).Offset(-1, 0) And Range("P" & i) = "OSP" Then

            If Left(Range("AP" & i), 1) Like "[a-z, A-Z]" Then
                
            Else
                Rows(i).Select
                Selection.Delete Shift:=xlUp
                i = i - 1
            End If
        End If
    Next

    Sheets("OSP轉正航單據").Range("B1").Formula = "=500-COUNTBLANK(RD訂單單據轉出!A1:A500)+1"

    Application.CutCopyMode = False
    Sheets("RD訂單單據轉出").Copy
    Sheets("RD訂單單據轉出").Select
    Sheets("RD訂單單據轉出").Name = YEE & "RD訂單單據轉出"
    ChDir "\\YEAWEN\files-server\06_資材\01_生管\航電每日資訊\航電訂單銷貨轉正航\航電訂單_轉正航單據"
    ActiveWorkbook.SaveAs Filename:= _
        "\\YEAWEN\files-server\06_資材\01_生管\航電每日資訊\航電訂單銷貨轉正航\航電訂單_轉正航單據\" & YEE & "轉正航訂單單據.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close

    Sheets("DATA").Select
    Range("H1").Select

    複製訂單到MARS表
    
End Sub


