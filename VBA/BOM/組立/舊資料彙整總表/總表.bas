Attribute VB_Name = "總表"
Sub 總表()

    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = "客戶主檔"
    
    '----------客戶----------
    Worksheets("test").Select
    Range("B5", Range("B" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("A1").PasteSpecial xlPasteValues
    '----------客戶----------
    
    
    '----------狀態----------
    Range("B1") = "狀態"
    '----------狀態----------
    
    
    '----------圖示----------
    Range("C1") = "圖示"
    '----------圖示----------
    
    
    '----------產品別----------
    Range("D1") = "產品別"
    '----------產品別----------
    
    
    '----------機種----------
    Worksheets("test").Select
    Range("C5", Range("C" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("E1").PasteSpecial xlPasteValues
    '----------機種----------
    
    
    '----------階層次序----------
    Worksheets("test").Select
    Range("D5", Range("D" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("F1").PasteSpecial xlPasteValues
    '----------階層次序----------
    
    
    '----------成品料號----------
    Worksheets("test").Select
    Range("E5", Range("E" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("G1").PasteSpecial xlPasteValues
    Range("G1") = "成品料號"
    '----------成品料號----------
    
    
    '----------Lever1----------
    Worksheets("test").Select
    Range("F5", Range("F" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("H1").PasteSpecial xlPasteValues
    '----------Lever1----------
    
    
    '----------Lever2----------
    Worksheets("test").Select
    Range("G5", Range("G" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("I1").PasteSpecial xlPasteValues
    '----------Lever2----------
    
    
    '----------Lever3----------
    Worksheets("test").Select
    Range("H5", Range("H" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("J1").PasteSpecial xlPasteValues
    '----------Lever3----------
    
    '----------Lever4----------
    Range("K1") = "Lever4"
    '----------Lever4----------
    
    
    '----------Lever5----------
    Range("L1") = "Lever5"
    '----------Lever5----------
    
    
    '----------廠商----------
    Worksheets("test").Select
    Range("K5", Range("K" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("M1").PasteSpecial xlPasteValues
    Range("M1") = "廠商"
    '----------廠商----------
    
    
    '----------用量----------
    Worksheets("test").Select
    Range("L5", Range("L" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("N1").PasteSpecial xlPasteValues
    Range("N1") = "用量"
    '----------用量----------
    
    
    '----------標準損耗----------
    Worksheets("test").Select
    Range("M5", Range("M" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("O1").PasteSpecial xlPasteValues
    Range("O1") = "標準損耗"
    '----------標準損耗----------
    
    
    '----------品名----------
    Worksheets("test").Select
    Range("I5", Range("I" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("P1").PasteSpecial xlPasteValues
    Range("P1") = "品名"
    '----------品名----------
    
    
    '----------規格----------
    Worksheets("test").Select
    Range("J5", Range("J" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("Q1").PasteSpecial xlPasteValues
    Range("Q1") = "規格"
    '----------規格----------
    
    
    '----------成品重量----------
    Worksheets("test").Select
    Range("W5", Range("W" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("R1").PasteSpecial xlPasteValues
    Range("R1") = "成品重量"
    
    Range("T2").Select
    ActiveCell.Formula = "=LEFT(R2,FIND(""/"",R2)-1)"
    Range("T2").Select
    
    
    Dim lrow As Long
    lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    Selection.AutoFill Destination:=Range("T2:T" & lrow)
    
        '----------包裝數量----------
        Worksheets("客戶主檔").Select
        Range("S1") = "包裝數量"
        
        Range("U2").Select
        ActiveCell.Formula = "=LEFT(R2, FIND(""_"", R2)) & MID(R2,FIND(""/"",R2)+1,5)"
        Range("U2").Select
        Selection.AutoFill Destination:=Range("U2:U" & lrow)
        
        Range("U2", Range("U" & Range("U65536").End(xlUp).Row)).Copy
        Range("S2").PasteSpecial xlPasteValues
        Range("U2", Range("U" & Range("U65536").End(xlUp).Row)).ClearContents
        '----------包裝數量----------
    
    Range("T2", Range("T" & Range("T65536").End(xlUp).Row)).Copy
    Range("R2").PasteSpecial xlPasteValues
    Range("T2", Range("T" & Range("T65536").End(xlUp).Row)).ClearContents
    '----------成品重量----------
    
    
    '----------週期(報價工時)----------
    Worksheets("test").Select
    Range("X5", Range("X" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("T1").PasteSpecial xlPasteValues
    Range("T1") = "週期(報價工時)"
    '----------週期(報價工時)----------
    


    '----------實際工時----------
    Range("U1") = "實際工時"
    '----------實際工時----------
    
    
    '----------實際工站----------
    Range("V1") = "實際工站"
    '----------實際工站----------
    
    
    '----------加工項目1----------
    Range("W1") = "加工項目1"
    '----------加工項目1----------


    '----------加工項目2----------
    Range("X1") = "加工項目2"
    '----------加工項目2----------
    
    
    '----------加工項目3----------
    Range("Y1") = "加工項目3"
    '----------加工項目3----------
    
    
    '----------加工項目4----------
    Range("Z1") = "加工項目4"
    '----------加工項目4----------
    
    
    '----------實際損耗----------
    Range("AA1") = "實際損耗"
    '----------實際損耗----------
    
    
    '----------機構工程師----------
    Range("AB1") = "機構工程師"
    '----------機構工程師----------
    
    
    '----------製程工程師----------
    Range("AC1") = "製程工程師"
    '----------製程工程師----------
    
    
    '----------電子工程師----------
    Range("AD1") = "電子工程師"
    '----------電子工程師----------
    
    
    '----------FAID----------
    Range("AE1") = "FAID"
    '----------FAID----------
    
    
    '----------版本----------
    Range("AF1") = "版本"
    '----------版本----------
    
    
    '----------備註----------
    Worksheets("test").Select
    Range("Y5", Range("Y" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("AG1").PasteSpecial xlPasteValues
    Range("AG1") = "備註"
    '----------備註----------
    
    
    '----------產品履歷----------
    Range("AH1") = "產品履歷"
    '----------產品履歷----------
    
    
    '----------BOM----------
    Range("AI1") = "BOM"
    '----------BOM----------
    
    
    '----------成品圖----------
    Range("AJ1") = "成品圖"
    '----------成品圖----------
    
    
    '----------FA----------
    Range("AK1") = "FA"
    '----------FA----------
    
    
    
    '----------零件圖1----------
    Range("AL1") = "零件圖1"
    '----------零件圖1----------
    
    
    
    '----------日期版本1----------
    Range("AM1") = "日期版本1"
    '----------日期版本1----------
    
    
    '----------零件圖2----------
    Range("AN1") = "零件圖2"
    '----------零件圖2----------
    
    
    '----------日期版本2----------
    Range("AO1") = "日期版本2"
    '----------日期版本2----------
    
    
    '----------零件圖3----------
    Range("AP1") = "零件圖3"
    '----------零件圖3----------
    
    
    '----------日期版本3----------
    Range("AQ1") = "日期版本3"
    '----------日期版本3----------
    
    
    '----------零件圖4----------
    Range("AR1") = "零件圖4"
    '----------零件圖4----------
    
    
    '----------日期版本4----------
    Range("AS1") = "日期版本4"
    '----------日期版本4----------
    
    
    '----------零件圖5----------
    Range("AT1") = "零件圖5"
    '----------零件圖5----------
    
    
    '----------日期版本5----------
    Range("AU1") = "日期版本5"
    '----------日期版本5----------
    
    
    '----------零件圖6----------
    Range("AV1") = "零件圖6"
    '----------零件圖6----------
    
    
    '----------日期版本6----------
    Range("AW1") = "日期版本6"
    '----------日期版本6----------
    
    
    '----------零件圖7----------
    Range("AX1") = "零件圖7"
    '----------零件圖7----------
    
    
    '----------日期版本7----------
    Range("AY1") = "日期版本7"
    '----------日期版本7----------
    
    
    '----------零件圖8----------
    Range("AZ1") = "零件圖8"
    '----------零件圖8----------
    
    
    '----------日期版本8----------
    Range("BA1") = "日期版本8"
    '----------日期版本8----------
    
    
    '----------零件圖9----------
    Range("BB1") = "零件圖9"
    '----------零件圖9----------
    
    
    '----------日期版本9----------
    Range("BC1") = "日期版本9"
    '----------日期版本9----------
    
    
    '----------零件圖10----------
    Range("BD1") = "零件圖10"
    '----------零件圖10----------
    
    
    '----------日期版本10----------
    Range("BE1") = "日期版本10"
    '----------日期版本10----------
    
    
    '----------零件圖11----------
    Range("BF1") = "零件圖11"
    '----------零件圖11----------
    
    
    '----------日期版本11----------
    Range("BG1") = "日期版本11"
    '----------日期版本11----------
    
    
    
    '----------零件圖12----------
    Range("BH1") = "零件圖12"
    '----------零件圖12----------
    
    
    '----------日期版本12----------
    Range("BI1") = "日期版本12"
    '----------日期版本12----------
    
    
    '----------零件圖13----------
    Range("BJ1") = "零件圖13"
    '----------零件圖13----------
    
    
    '----------日期版本13----------
    Range("BK1") = "日期版本13"
    '----------日期版本13----------
    
    
    '----------零件圖14----------
    Range("BL1") = "零件圖14"
    '----------零件圖14----------
    
    
    '----------日期版本14----------
    Range("BM1") = "日期版本14"
    '----------日期版本14----------
    
    
    '----------零件圖15----------
    Range("BN1") = "零件圖15"
    '----------零件圖15----------
    
    
    '----------日期版本15----------
    Range("BO1") = "日期版本15"
    '----------日期版本15----------
    
    
    '----------零件圖16----------
    Range("BP1") = "零件圖16"
    '----------零件圖16----------
    
    
    '----------日期版本16----------
    Range("BQ1") = "日期版本16"
    '----------日期版本16----------
    
    
    '----------零件圖17----------
    Range("BR1") = "零件圖17"
    '----------零件圖17----------
    
    
    '----------日期版本17----------
    Range("BS1") = "日期版本17"
    '----------日期版本17----------
    
    
    '----------零件圖18----------
    Range("BT1") = "零件圖18"
    '----------零件圖18----------
    
    
    '----------日期版本18----------
    Range("BU1") = "日期版本18"
    '----------日期版本18----------
    
    
    '----------零件圖19----------
    Range("BV1") = "零件圖19"
    '----------零件圖19----------
    
    
    '----------日期版本19----------
    Range("BW1") = "日期版本19"
    '----------日期版本19----------
    
    
    '----------零件圖20----------
    Range("BX1") = "零件圖20"
    '----------零件圖20----------
    
    
    '----------日期版本20----------
    Range("BY1") = "日期版本20"
    '----------日期版本20----------
    
    
    '----------零件圖21----------
    Range("BZ1") = "零件圖21"
    '----------零件圖21----------
    
    
    '----------日期版本21----------
    Range("CA1") = "日期版本21"
    '----------日期版本21----------
    
    
    '----------零件圖22----------
    Range("CB1") = "零件圖22"
    '----------零件圖22----------
    
    
    '----------日期版本22----------
    Range("CC1") = "日期版本22"
    '----------日期版本22----------
    
    
    '----------零件圖23----------
    Range("CD1") = "零件圖23"
    '----------零件圖23----------
    
    
    '----------日期版本23----------
    Range("CE1") = "日期版本23"
    '----------日期版本23----------
    
    '----------零件圖24----------
    Range("CF1") = "零件圖24"
    '----------零件圖24----------
    
    
    '----------日期版本24----------
    Range("CG1") = "日期版本24"
    '----------日期版本24----------
    
    '----------零件圖25----------
    Range("CH1") = "零件圖25"
    '----------零件圖25----------
    
    
    '----------日期版本25----------
    Range("CI1") = "日期版本25"
    '----------日期版本25----------
    
    
    '----------零件圖26----------
    Range("CJ1") = "零件圖26"
    '----------零件圖26----------
    
    
    '----------日期版本26----------
    Range("CK1") = "日期版本26"
    '----------日期版本26----------
    
    
    '----------零件圖27----------
    Range("CL1") = "零件圖27"
    '----------零件圖27----------
    
    
    '----------日期版本27----------
    Range("CM1") = "日期版本27"
    '----------日期版本27----------
    
    
    '----------零件圖28----------
    Range("CN1") = "零件圖28"
    '----------零件圖28----------
    
    
    '----------日期版本28----------
    Range("CO1") = "日期版本28"
    '----------日期版本28----------
    
    
    '----------零件圖29----------
    Range("CP1") = "零件圖29"
    '----------零件圖29----------
    
    
    '----------日期版本29----------
    Range("CQ1") = "日期版本29"
    '----------日期版本29----------
    
    
    '----------零件圖30----------
    Range("CR1") = "零件圖30"
    '----------零件圖30----------
    
    
    '----------日期版本30----------
    Range("CS1") = "日期版本30"
    '----------日期版本30----------
    
    
    '----------廠內備註----------
    Range("CT1") = "廠內備註"
    '----------廠內備註----------
    
    
    '----------版本1----------
    Worksheets("test").Select
    Range("Z5", Range("Z" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("CU1").PasteSpecial xlPasteValues
    Range("CU1") = "版本1"
    '----------版本1----------
    
    
    '----------修訂日期1----------
    Worksheets("test").Select
    Range("AA5", Range("AA" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("CV1").PasteSpecial xlPasteValues
    Range("CV1") = "修訂日期1"
    '----------修訂日期1----------
    
    
    
    '----------變更記錄1----------
    Worksheets("test").Select
    Range("AB5", Range("AB" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("CW1").PasteSpecial xlPasteValues
    Range("CW1") = "變更記錄1"
    '----------變更記錄1----------
    
    
    '----------核準1----------
    Worksheets("test").Select
    Range("AC5", Range("AC" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("CX1").PasteSpecial xlPasteValues
    Range("CX1") = "核準1"
    '----------核準1----------
    
    
    '----------審核1----------
    Worksheets("test").Select
    Range("AD5", Range("AD" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("CY1").PasteSpecial xlPasteValues
    Range("CY1") = "審核1"
    '----------審核1----------
    
    
    '----------製表1----------
    Worksheets("test").Select
    Range("AE5", Range("AE" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("CZ1").PasteSpecial xlPasteValues
    Range("CZ1") = "製表1"
    '----------製表1----------
    
    
    
    '----------版本2----------
    Worksheets("test").Select
    Range("AF5", Range("AF" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DA1").PasteSpecial xlPasteValues
    Range("DA1") = "版本2"
    '----------版本2----------
    
    
    '----------修訂日期2----------
    Worksheets("test").Select
    Range("AG5", Range("AG" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DB1").PasteSpecial xlPasteValues
    Range("DB1") = "修訂日期2"
    '----------修訂日期2----------
    
    
    
    '----------變更記錄2----------
    Worksheets("test").Select
    Range("AH5", Range("AH" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DC1").PasteSpecial xlPasteValues
    Range("DC1") = "變更記錄2"
    '----------變更記錄2----------
    
    
    '----------核準2----------
    Worksheets("test").Select
    Range("AI5", Range("AI" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DD1").PasteSpecial xlPasteValues
    Range("DD1") = "核準2"
    '----------核準2----------
    
    
    '----------審核2----------
    Worksheets("test").Select
    Range("AJ5", Range("AJ" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DE1").PasteSpecial xlPasteValues
    Range("DE1") = "審核2"
    '----------審核2----------
    
    
    '----------製表2----------
    Worksheets("test").Select
    Range("AK5", Range("AK" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DF1").PasteSpecial xlPasteValues
    Range("DF1") = "製表2"
    '----------製表2----------
    
    
    '----------版本3----------
    Worksheets("test").Select
    Range("AL5", Range("AL" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DG1").PasteSpecial xlPasteValues
    Range("DG1") = "版本3"
    '----------版本3----------
    
    
    '----------修訂日期3----------
    Worksheets("test").Select
    Range("AM5", Range("AM" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DH1").PasteSpecial xlPasteValues
    Range("DH1") = "修訂日期3"
    '----------修訂日期3----------
    
    
    
    '----------變更記錄3----------
    Worksheets("test").Select
    Range("AN5", Range("AN" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DI1").PasteSpecial xlPasteValues
    Range("DI1") = "變更記錄3"
    '----------變更記錄3----------
    
    
    '----------核準3----------
    Worksheets("test").Select
    Range("AO5", Range("AO" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DJ1").PasteSpecial xlPasteValues
    Range("DJ1") = "核準3"
    '----------核準3----------
    
    
    '----------審核3----------
    Worksheets("test").Select
    Range("AP5", Range("AP" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DK1").PasteSpecial xlPasteValues
    Range("DK1") = "審核3"
    '----------審核3----------
    
    
    '----------製表3----------
    Worksheets("test").Select
    Range("AQ5", Range("AQ" & Range("B65536").End(xlUp).Row)).Copy
    
    Worksheets("客戶主檔").Select
    Range("DL1").PasteSpecial xlPasteValues
    Range("DL1") = "製表3"
    '----------製表3----------
    
    
    
    '----------版本4----------
    Range("DM1") = "版本4"
    '----------版本4----------
    
    
    '----------修訂日期4----------
    Range("DN1") = "修訂日期4"
    '----------修訂日期4----------
    
    
    '----------變更記錄4----------
    Range("DO1") = "變更記錄4"
    '----------變更記錄4----------
    
    
    '----------核準4----------
    Range("DP1") = "核準4"
    '----------核準4----------
    
    
    '----------審核4----------
    Range("DQ1") = "審核4"
    '----------審核4----------
    
    
    '----------製表4----------
    Range("DR1") = "製表4"
    '----------製表4----------
    
    
    '----------版本5----------
    Range("DS1") = "版本5"
    '----------版本5----------
    
    
    '----------修訂日期5----------
    Range("DT1") = "修訂日期5"
    '----------修訂日期5----------
    
    
    '----------變更記錄5----------
    Range("DU1") = "變更記錄5"
    '----------變更記錄5----------
    
    
    '----------核準5----------
    Range("DV1") = "核準5"
    '----------核準5----------
    
    
    '----------審核5----------
    Range("DW1") = "審核5"
    '----------審核5----------
    
    
    '----------製表5----------
    Range("DX1") = "製表5"
    '----------製表5----------
    
    
    
    '----------成品----------
    Range("DY1") = "成品"
    '----------成品----------
    
    
    '----------零件1----------
    Range("DZ1") = "零件1"
    '----------零件1----------
    
    
    '----------日期版本1----------
    Range("EA1") = "日期版本1"
    '----------日期版本1----------
    
    
    
    '----------零件2----------
    Range("EB1") = "零件2"
    '----------零件2----------
    
    
    '----------日期版本2----------
    Range("EC1") = "日期版本2"
    '----------日期版本2----------
    
    
    '----------零件3----------
    Range("ED1") = "零件3"
    '----------零件3----------
    
    
    '----------日期版本3----------
    Range("EE1") = "日期版本3"
    '----------日期版本3----------
    
    
    '----------零件4----------
    Range("EF1") = "零件4"
    '----------零件4----------
    
    
    '----------日期版本4----------
    Range("EG1") = "日期版本4"
    '----------日期版本4----------
    

    '----------零件5----------
    Range("EH1") = "零件5"
    '----------零件5----------
    
    
    '----------日期版本5----------
    Range("EI1") = "日期版本5"
    '----------日期版本5----------
    

    '----------零件6----------
    Range("EJ1") = "零件6"
    '----------零件6----------
    
    
    '----------日期版本6----------
    Range("EK1") = "日期版本6"
    '----------日期版本6----------
    
    
    '----------零件7----------
    Range("EL1") = "零件7"
    '----------零件7----------
    
    
    '----------日期版本7----------
    Range("EM1") = "日期版本7"
    '----------日期版本7----------
    
    
    '----------零件8----------
    Range("EN1") = "零件8"
    '----------零件8----------
    
    
    '----------日期版本8----------
    Range("EO1") = "日期版本8"
    '----------日期版本8----------
    

    '----------零件9----------
    Range("EP1") = "零件9"
    '----------零件9----------
    
    
    '----------日期版本9----------
    Range("EQ1") = "日期版本9"
    '----------日期版本9----------
    

    '----------零件10----------
    Range("ER1") = "零件10"
    '----------零件10----------
    
    
    '----------日期版本10----------
    Range("ES1") = "日期版本10"
    '----------日期版本10----------
    

    '----------零件11----------
    Range("ET1") = "零件11"
    '----------零件11----------
    
    
    '----------日期版本11----------
    Range("EU1") = "日期版本11"
    '----------日期版本11----------
    

    '----------零件12----------
    Range("EV1") = "零件12"
    '----------零件12----------
    
    
    '----------日期版本12----------
    Range("EW1") = "日期版本12"
    '----------日期版本12----------
    

    '----------零件13----------
    Range("EX1") = "零件13"
    '----------零件13----------
    
    
    '----------日期版本13----------
    Range("EY1") = "日期版本13"
    '----------日期版本13----------
    

    '----------零件14----------
    Range("EZ1") = "零件14"
    '----------零件14----------
    
    
    '----------日期版本14----------
    Range("FA1") = "日期版本14"
    '----------日期版本14----------
    

    '----------零件15----------
    Range("FB1") = "零件15"
    '----------零件15----------
    
    
    '----------日期版本15----------
    Range("FC1") = "日期版本15"
    '----------日期版本15----------
    

    '----------零件16----------
    Range("FD1") = "零件16"
    '----------零件16----------
    
    
    '----------日期版本16----------
    Range("FE1") = "日期版本16"
    '----------日期版本16----------
    

    '----------零件17----------
    Range("FF1") = "零件17"
    '----------零件17----------
    
    
    '----------日期版本17----------
    Range("FG1") = "日期版本17"
    '----------日期版本17----------
    

    '----------零件18----------
    Range("FH1") = "零件18"
    '----------零件18----------
    
    
    '----------日期版本18----------
    Range("FI1") = "日期版本18"
    '----------日期版本18----------
    

    '----------零件19----------
    Range("FJ1") = "零件19"
    '----------零件19----------
    
    
    '----------日期版本19----------
    Range("FK1") = "日期版本19"
    '----------日期版本19----------
    

    '----------零件20----------
    Range("FL1") = "零件20"
    '----------零件20----------
    
    
    '----------日期版本20----------
    Range("FM1") = "日期版本20"
    '----------日期版本20----------
    

    '----------零件21----------
    Range("FN1") = "零件21"
    '----------零件21----------
    
    
    '----------日期版本21----------
    Range("FO1") = "日期版本21"
    '----------日期版本21----------
    

    '----------零件22----------
    Range("FP1") = "零件22"
    '----------零件22----------
    
    
    '----------日期版本22----------
    Range("FQ1") = "日期版本22"
    '----------日期版本22----------
    

    '----------零件23----------
    Range("FR1") = "零件23"
    '----------零件23----------
    
    
    '----------日期版本23----------
    Range("FS1") = "日期版本23"
    '----------日期版本23----------
    

    '----------零件24----------
    Range("FT1") = "零件24"
    '----------零件24----------
    
    
    '----------日期版本24----------
    Range("FU1") = "日期版本24"
    '----------日期版本24----------
    

    '----------零件25----------
    Range("FV1") = "零件25"
    '----------零件25----------
    
    
    '----------日期版本25----------
    Range("FW1") = "日期版本25"
    '----------日期版本25----------
    

    '----------零件26----------
    Range("FX1") = "零件26"
    '----------零件26----------
    
    
    '----------日期版本26----------
    Range("FY1") = "日期版本26"
    '----------日期版本26----------
    

    '----------零件27----------
    Range("FZ1") = "零件27"
    '----------零件27----------
    
    
    '----------日期版本27----------
    Range("GA1") = "日期版本27"
    '----------日期版本27----------
    

    '----------零件28----------
    Range("GB1") = "零件28"
    '----------零件28----------
    
    
    '----------日期版本28----------
    Range("GC1") = "日期版本28"
    '----------日期版本28----------
    

    '----------零件29----------
    Range("GD1") = "零件29"
    '----------零件29----------
    
    
    '----------日期版本29----------
    Range("GE1") = "日期版本29"
    '----------日期版本29----------
    

    '----------零件30----------
    Range("GF1") = "零件30"
    '----------零件30----------
    
    
    '----------日期版本30----------
    Range("GG1") = "日期版本30"
    '----------日期版本30----------
    
    
    '----------------調整日期格式----------------
    Range("CV:CV, DB:DB, DH:DH, DN:DN, DT:DT").Select
    Selection.NumberFormatLocal = "yyyy/mm/dd"
    '----------------調整日期格式----------------
    

    Application.CutCopyMode = False

End Sub
