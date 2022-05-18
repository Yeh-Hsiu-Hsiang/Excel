
Sub GeneralTable_首件()
Attribute GeneralTable_首件.VB_ProcData.VB_Invoke_Func = " \n14"

Dim ActWb As String, i, j, k As Long
    
ActWb = ActiveWorkbook.Name


'-----------------匯出資料總表整理-----------------
Range("A:A, C:F, W:W, EX:EX, LL:LL").Select
Selection.Copy

Worksheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False


Workbooks(ActWb).Worksheets(1).Activate

'-------------綜合 判定-------------
Range("EW:EW , LK:LK").Select
Selection.Copy

Worksheets(2).Activate

i = 1
Do While True
    If ActiveSheet.Cells(1, i).Value = "" Then
        ActiveSheet.Cells(1, i).Select
        Exit Do
    End If
    i = i + 1
Loop

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
'-------------綜合 判定-------------


'-------------檢驗異常備註-------------
Workbooks(ActWb).Worksheets(1).Activate
Range("EZ:EZ , LN:LN").Select
Selection.Copy

Worksheets(2).Activate

i = 1
Do While True
    If ActiveSheet.Cells(1, i).Value = "" Then
        ActiveSheet.Cells(1, i).Select
        Exit Do
    End If
    i = i + 1
Loop

Selection.PasteSpecial xlPasteValues
Application.CutCopyMode = False
'-------------檢驗異常備註-------------


'-----------------匯出資料總表整理-----------------

Dim lrow As Long
lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

'-----------------總表 首件-----------------
Columns("B:B").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("B2").Select
ActiveCell.FormulaR1C1 = "首件"
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B" & lrow)


Columns("A:A").Select
Selection.NumberFormatLocal = "yyyy/mm/dd"
Application.WindowState = xlNormal

Columns("G:G").Select
Selection.NumberFormatLocal = "yyyy/mm/dd"
Application.WindowState = xlNormal

'-----------------總表 首件-----------------



'-----------------總表 檢驗員-----------------
Columns("H:H").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("H1") = "檢驗員"
Range("H2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(I2=J2, I2, I2 & "" "" & J2)"
Range("H2").Select
Selection.AutoFill Destination:=Range("H2:H" & lrow)
'-----------------總表 檢驗員-----------------


'-----------------總表 製造數-----------------
Range("O1") = "製造數"
Range("O2") = 1
Range("O2").Select
Selection.AutoFill Destination:=Range("O2:O" & lrow)
'-----------------總表 製造數-----------------


'-----------------總表 抽驗數-----------------
Range("P1") = "抽驗數"
Range("P2") = 1
Range("P2").Select
ActiveCell.Formula = "=IF(AND(O2>=2, O2<=544), 32, IF(AND(O2>=545, O2<=960), 40,  IF(AND(O2>=961, O2<=1632), 48,  IF(AND(O2>=1633, O2<=3072), 64,  IF(O2>=3073, 80, 1)))))"
Selection.AutoFill Destination:=Range("P2:P" & lrow)
'-----------------總表 抽驗數-----------------


'-----------------總表 不良數-----------------
Range("Q1") = "不良數"
Range("Q2") = 0
Range("Q2").Select
Selection.AutoFill Destination:=Range("Q2:Q" & lrow)
'-----------------總表 不良數-----------------


'-----------------總表 抽驗不良率-----------------
Range("R1") = "抽驗不良率"
Range("R2").Formula = "=IFERROR(Q2/P2, 0)"
Range("R2").Select
Selection.AutoFill Destination:=Range("R2:R" & lrow)
'-----------------總表 抽驗不良率-----------------



'-----------------總表 批不良率-----------------
Range("S1") = "批不良率"
Range("S2").Formula = "=IFERROR(Q2/O2, 0)"
Range("S2").Select
Selection.AutoFill Destination:=Range("S2:S" & lrow)
'-----------------總表 批不良率-----------------



'-----------------總表 綜合判定-----------------
Columns("K:K").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("K1") = "綜合判定"
Range("K2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(FIND(""可生產"",L2)>4, ""NG"", ""OK"")"
Range("K2").Select
Selection.AutoFill Destination:=Range("K2:K" & lrow)
'-----------------總表 綜合判定-----------------



'-----------------總表 檢驗異常備註-----------------
Columns("N:N").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("N1") = "檢驗異常備註"
Range("N2").Select
Application.CutCopyMode = False
ActiveCell.Formula = "=IF(O2="""", """", IF(P2="""", O2, O2 & ""。  "" & P2))"
Range("N2").Select
Selection.AutoFill Destination:=Range("N2:N" & lrow)
'-----------------總表 檢驗異常備註-----------------



'-----------------總表 NG數-----------------
Range("V1") = "NG數"
Range("V2").Formula = "=COUNTIF(K2, ""NG"")"
Range("V2").Select
Selection.AutoFill Destination:=Range("V2:V" & lrow)
'-----------------總表 NG數-----------------



'-----------------總表 NG數資料-----------------
For k = 2 To 5000
    
    If Range("K" & k) = "NG" Then
        
        If Range("A" & k) = Range("A" & k).Offset(-1, 0) And Range("C" & k) = Range("C" & k).Offset(-1, 0) Then
            k = k + 1
        Else
            For m = 1 To Range("V" & k)

                Range("A" & k & ":V" & k).Select
                Selection.Copy

                Range("A" & k & ":V" & k).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
            Next m

            Range("K" & k) = "OK"
            Range("S" & k) = 0
        End If
    End If
Next
'-----------------總表 NG數資料-----------------



'-----------------匯出資料總表整理-----------------


'複製資料匯出總表 準備貼到品保 IPQC 總表
Range("B2", ActiveSheet.Range("B" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy


'-----------------D 欄 IPQC-----------------
Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

j = 6
Do While True
    If Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Cells(j, "D").Value = "" Then
        Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Cells(j, "D").Select
        Exit Do
    End If
    j = j + 1
Loop



ActiveSheet.Range("D" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------D 欄 IPQC-----------------



'-----------------檢驗日期-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("A2", ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("E" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗日期-----------------



'-----------------檢驗員-----------------
Workbooks(ActWb).Worksheets(2).Activate

ActiveSheet.Range("H2", ActiveSheet.Range("H" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("F" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------檢驗員-----------------



'-----------------製令單號-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("C2", ActiveSheet.Range("C" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("H" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令單號-----------------



'-----------------製令日期-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("G2", ActiveSheet.Range("G" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("I" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製令日期-----------------



'-----------------客戶-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("D2", ActiveSheet.Range("D" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("J" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------客戶-----------------



'-----------------機種-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("E2", ActiveSheet.Range("E" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("K" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------機種-----------------



'-----------------品名-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("F2", ActiveSheet.Range("F" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("L" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------品名-----------------



'-----------------製造數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("Q2", ActiveSheet.Range("Q" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("R" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------製造數-----------------



'-----------------抽驗數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("R2", ActiveSheet.Range("R" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("S" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------抽驗數-----------------



'-----------------不良數-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("S2", ActiveSheet.Range("S" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("T" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------不良數-----------------



'-----------------抽驗不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("T2", ActiveSheet.Range("T" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("V" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------抽驗不良率-----------------



'-----------------批不良率-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("U2", ActiveSheet.Range("U" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("W" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------批不良率-----------------



'-----------------判定-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("K2", ActiveSheet.Range("K" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("U" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------判定-----------------



'-----------------備註1-----------------
Workbooks(ActWb).Worksheets(2).Activate
ActiveSheet.Range("N2", ActiveSheet.Range("N" & ActiveSheet.Rows.Count).End(xlUp)).Select
Selection.Copy

Workbooks("品保IPQC_FQC日報系統(組立20210305.xlsm").Worksheets("Q品質檢驗資料總表(加工)").Activate

ActiveSheet.Range("AC" & j).Select
Selection.PasteSpecial xlPasteValues
'-----------------備註1-----------------

End Sub

