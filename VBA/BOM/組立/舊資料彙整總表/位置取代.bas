Attribute VB_Name = "位置取代"
Sub 位置取代()

Cells.Replace "$C", "01", xlPart    '階層次序 C
Cells.Replace "$D", "02", xlPart    '組立品番 D
Cells.Replace "$E", "03", xlPart    'Lever1 E
Cells.Replace "$F", "04", xlPart    'Lever2 F
Cells.Replace "$G", "05", xlPart    'Lever3 G



Cells.Replace "$H", "08", xlPart    '品名 H(06) -> J(08)
Cells.Replace "$I", "09", xlPart    '規格 I(07) -> K(09)
Cells.Replace "$J", "10", xlPart    '廠商 J(08) -> L(10)
Cells.Replace "$K", "11", xlPart    '用量 K(09) -> M(11)
Cells.Replace "$L", "12", xlPart    '標準損耗 L(10) -> N(12)


'
'Cells.Replace "$M", "11", xlPart
'Cells.Replace "$N", "12", xlPart
'Cells.Replace "$O", "13", xlPart
'Cells.Replace "$Q", "14", xlPart
'Cells.Replace "$S", "15", xlPart
'Cells.Replace "$U", "16", xlPart



Columns("R").Replace "$V", "17", xlPart     '單重 V -> W(17)
Columns("S").Replace "$V", "18", xlPart     '包裝數 V -> X(18)
Cells.Replace "$X", "19", xlPart            '週期 X(18) -> Y(19)
Cells.Replace "$Y", "20", xlPart            '備註 Y(19) -> Z(20)

End Sub
