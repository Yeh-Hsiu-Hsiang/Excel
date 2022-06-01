Attribute VB_Name = "料號更改"

Sub 料號更改()
  Sheets("客戶主檔").Cells.Replace _
  Sheets("加工項目查詢及料號更改").[R3], _
  Sheets("加工項目查詢及料號更改").[S3], xlWhole
  
  MsgBox ("修改完成")
End Sub
