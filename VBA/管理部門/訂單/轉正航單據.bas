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
        
        
    '---------------------自動取代正航版本---------------------
    Sheets("單價").Select
    Range("AT1") = "正航料號"
    Range("AU1") = "正航版本"
    Range("AV1") = "訂單料號"
    Range("AW1") = "訂單版本"
    Range("AX1") = "OSP料號"
    Range("AY1") = "OSP版本"
    Range("AZ1") = "訂單依正航版本為主"
    Range("BA1") = "OSP依正航版本為主"
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    Range("AT2").Select
    ActiveCell.Formula = "=I2"
    Selection.AutoFill Destination:=Range("AT2:AT" & lrow)
    
    
    Range("AU2").Select
    ActiveCell.Formula = "=J2"
    Selection.AutoFill Destination:=Range("AU2:AU" & lrow)
    
    Range("AV2").Select
    ActiveCell.Formula = "=IF(N2="""","""",LEFT(N2,FIND(""#"",N2,1)-1))"
    Selection.AutoFill Destination:=Range("AV2:AV" & lrow)
    
    Range("AW2").Select
    ActiveCell.Formula = "=IF(N2="""","""",MID(N2,FIND(""#"",N2,1)+1,5))"
    Selection.AutoFill Destination:=Range("AW2:AW" & lrow)
    
    Range("AX2").Select
    ActiveCell.Formula = "=IF(OSP!C5="""","""",OSP!C5)"
    Selection.AutoFill Destination:=Range("AX2:AX" & lrow)
    
    Range("AY2").Select
    ActiveCell.Formula = "=IF(OSP!C5="""","""",OSP!D5)"
    Selection.AutoFill Destination:=Range("AY2:AY" & lrow)
    
    For j = 2 To Range("AX65536").End(xlUp).Row
        If Range("AX" & j) = "" And Range("AY" & j) = "" And Range("AX" & j).Offset(1, 0) <> "" Then
            Range("AX" & j & ":AY" & j).Select
            Selection.Delete Shift:=xlUp
            j = j - 1
        End If
    Next
    
    Range("AZ2").Select
    ActiveCell.Formula = "=SUBSTITUTE(IFERROR(INDEX(A:A,MATCH(AV2,AT:AT,0),1), IF(AV2="""","""", AV2&""#""&AW2)),""#0"",""#O"",1)"
    Selection.AutoFill Destination:=Range("AZ2:AZ" & lrow)
    
    Range("BA2").Select
    ActiveCell.Formula = "=SUBSTITUTE(IFERROR(INDEX(A:A,MATCH(AX2,AT:AT,0),1), IF(AX2="""","""", AX2&""#""&AY2)),""#0"",""#O"",1)"
    Selection.AutoFill Destination:=Range("BA2:BA" & lrow)
    
    ActiveSheet.Range("AZ2", ActiveSheet.Range("AZ" & ActiveSheet.Rows.Count).End(xlUp)).Select
    Selection.Copy
    
    Worksheets("RD訂單單據轉出").Activate
    ActiveSheet.Range("AP2").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    Worksheets("單價").Activate
    ActiveSheet.Range("BA2", ActiveSheet.Range("BA" & ActiveSheet.Rows.Count).End(xlUp)).Select
    Selection.Copy

    Worksheets("OSP轉正航單據").Activate
    ActiveSheet.Range("AR2").Select
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    '---------------------自動取代正航版本---------------------

        
        
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
      
    整理單價明細
    '印單價差異
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
'    ChDir "\\YEAWEN\files-server\06_資材\01_生管\航電每日資訊\航電訂單銷貨轉正航\航電訂單_轉正航單據"
'    ActiveWorkbook.SaveAs Filename:= _
'        "\\YEAWEN\files-server\06_資材\01_生管\航電每日資訊\航電訂單銷貨轉正航\航電訂單_轉正航單據\" & YEE & "轉正航訂單單據.xlsx" _
'        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
'    ActiveWindow.Close
'
'    Sheets("DATA").Select
'    Range("H1").Select
'
'    複製訂單到MARS表
    
End Sub




