Sub time()
   Dim now_address As Range
   'Set now_address = ActiveCell    '設定目前儲存格位置（單一）
   Set now_address_Range = Selection   '設定目前選取儲存格位置（範圍）
   now_address.Value = "19:00" 
End Sub