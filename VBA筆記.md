# VBA 筆記

## 個人巨集位置
>   ```C:\Users\UserName\AppData\Roaming\Microsoft\Excel\XLSTART```

---
## 語法

>   ## Sub 副程式
>   ```VBA
>   Sub ProjectName()
>       ...
>   End Sub
>   ```
> 
>   ## Function 函數（傳回值時使用）
>   ```VBA
>   Function Add_int(ByRef a As Integer)
>       a = a + 2
>   End Function
> 
>   ---
> 
>   Sub CallAdd_int()
>       Dim num As integer 'num=0
>       Add_int(num)
>       MsgBox num 'display 0
>   End Sub
>   ```

---
## 訊息視窗：MsgBox
```VBA
MsgBox ("Hello, world!")
```

---
## 輸出至即時運算視窗：Debug.Print
```VBA
Debug.Print "s = " & s
```

---
## 清除即時運算視窗內容：SendKeys "^g^a{DEL}"
```VBA
Sub Del_list()
       SendKeys "^g^a{DEL}"
End Sub
```

[Excel VBA 除錯技巧：Debug.Print 與即時運算視窗使用教學](https://officeguide.cc/excel-vba-debug-immediate-window-tutorial/)

---
## 註解：```'```
```VBA
'MsgBox ("Hello, world!")
```

---
## 程式換行：```_```
```VBA
x = 1 + 2 + 3 + _
    4 + 5 + 6
```


---
## 文字換行顯示：```vbCrLf、Chr(10)```
```VBA
MsgBox "NOTICE:" & vbCrLf & "This is an Important Message!"
MsgBox "NOTICE:" & Chr(10) & "This is an Important Message!"
```

---
## Cells 儲存格
>   Cells(列,欄)  
>   **```ActiveCell '目前儲存格位置```**  
>   ```Cells(1,2)　'B1```  
>   ```Cells(1,"B") '=Cells("1","B")```  

---
## 清除儲存格
>   ```Range("A1:A2").ClearContents```  
>   ```Range("A1").Value = ""```

---
## Rows 列  
>   ```Rows(1) '=Rows("1")```  
>   ```Rows("1:3") '代表第一到第三列```   
>   ```Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row   '最後一列```   
>   Ex:  
>   ```
>   Range("D65536").End(xlUp).Row    '找到D欄最後一列
>   ```
>   
>   Ex：自動複製填滿某欄位直到最後一行
>   ```
>   Range("A1").Select
>   Dim lrow As Long
>   lrow = Cells(Cells.Rows.Count, "C").End(xlUp).Row
>   Selection.AutoFill Destination:=Range("A1:A" & lrow)
>   ```

---
## Columns 欄
>   ```Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column '最後一欄```  
>   ```Columns(4) '=Columns("D")```  
>   **如果要選取多欄在雙引號裡面要用英文字。**  
>   Ex:  C 欄到 D 欄  
>   ```VBA
>   Columns("C:D")
>   ```  
>   

---
## Range 區
>   Range 是 VBA 裡面最好用的選取方式  
>   支援單格、多格、單欄、多欄、單列及多列  
>   列的表示方式：把英文字改成數字  
>   * 單格：```Range("B1")```
>   * 多格：```Range("A1,B2,C3,D4")```  
>   * 單欄：```Range("A:A")```
>   * 多欄：```Range("B:B,E:E")```
> 
>   Range("A1","B2") 表示一區 (A1、B1、A2、B2) = Range(Cells("A1"),Cells("B2"))

---
## 取消選取
>   ```VBA
>   Application.CutCopyMode = False
>   ```

---
## 找到並選定最後一列的儲存格
>   ```VBA
>   i = 3
>   Do While True
>       If ActiveSheet.Cells(i, 1).Value = "" Then
>           ActiveSheet.Cells(i, 1).Select
>           Exit Do
>       End If
>       i = i + 1
>   Loop
>   ```

---
## 找到並選定最後一欄的儲存格
>   ```VBA
>   i = 3
>   Do While True
>       If ActiveSheet.Cells(1, i).Value = "" Then
>           ActiveSheet.Cells(1, i).Select
>           Exit Do
>       End If
>       i = i + 1
>   Loop
>   ```

---
## 活頁簿
>   ```VBA
>   Workbooks.Open "C:VBAdemo.xlsx" '開啟舊活頁簿
>
>   WorkBooks.Add   '開新活頁簿
>
>   Workbooks(1).Name   '活頁簿名稱
>
>   Workbooks.Count '目前活頁簿數量
>
>   Workbooks("book1.xlsm").Protect "password"  '加密活頁簿
>   Workbooks("book1.xlsm").Unprotect "password"    '解密活頁簿
>
>   Workbooks("demo").Save  '儲存活頁簿
>   WorkBook(1).Save    '第一本活頁簿儲存
>   Workbooks("demo").SaveAs "C:\Users\StevePC2\Downloads\new.xlsm"    '另存活頁簿
>   ActiveWorkbook.SaveAs Filename:= "C:\new.xlsx", WriteRes:= "password"   '另存成輸入密碼才能進行編輯的活頁簿
>
>   Workbooks("demo").Activate  'Activate可以指定當前活頁簿
>   ActiveWorkbook  '當前視窗活頁簿
>   WorkBook(2).Activate    '第二本活頁簿拉到當前視窗
>
>   Workbooks("demo").Close '關閉活頁簿
>   ActiveWorkbook.Close SaveChanges:=True  '關閉活頁簿並保存
>   Workbooks.Close '關閉所有活頁簿，但留下主視窗
>   Application.Quit    '關閉整個 Excel
>   ```

---
## 工作表
>   ```VBA
>   ActiveSheet '正在使用的工作表
>
>   Worksheet   '工作表
>   WorkSheets  '所有工作表
>
>   Worksheets.Add  '新增工作表
>   Worksheets.add().name = "test"  '新增一個 test 的工作表
>   Worksheets.add before:=Worksheets(2)    '在第二個工作表前新增工作表
>   Worksheets.add after:=Worksheets(1)     '在第一個工作表之後新增工作表
>   Worksheets.add count:=10    '新增十個工作表
>   Worksheets.Count    '現有工作表數量
>
>   '在最後一個工作表後新增一個叫 MySheet的工作表
>   Worksheets.add(after:=Worksheets(Worksheets.Count)).Name = "MySheet"    
>
>   Worksheets.add after:=Worksheets(Worksheets.Count), Count:=5    '在最後一個工作表後新增五個工作表
>
>   Sheets("Sheet1").Delete     '刪除工作表
>
>   Worksheets(1).Name = "新的工作表"   '改第一個工作表名稱
>
>   WorkSheets(1).Activate  '切換到第一個工作表
>   
>   Ex:
>   Worksheets("活頁簿1").Activate
>   ```

---
## 找尋特定工作表
>   ```VBA
>   Dim ws As Worksheet
>       
>   For Each ws In Worksheets   '讀取每個工作表
>       If ws.Name Like "*年" Then
>           ws.Activate
>           Exit For
>       End If
>   Next ws
>   ```



---
## 提示視窗

>   ```
>   Application.DisplayAlerts = False   '關閉提示視窗
>   Application.DisplayAlerts = True    '開啟提示視窗
>   ```

---
## 印表機
>   ```VBA
>   '指定 工作表1 的印表機為 Intermec PD43 (203 dpi)
>   Sheets(1).PrintOut ActivePrinter:="Intermec PD43 (203 dpi)"   
>   ```


---
## 給予值
>   Ex:  
>   A1、B1、A2、B2=100
>   ```VBA
>   Range("A1","B2").value=100
>   ```
>   ---
>   ```VBA
>   Range("A1") = Range("B1").Column    'A1的值等於B1的欄位 => A1=2  
> 
>   Range("C1") = Worksheets(1).Name    'C1的值等於第一個工作表的名字 => C1=工作表1  
> 
>   Range("D1") = Range("E1","E5").Count    'D1的值等於E1~E5的格數 => D1=5  
>   ``` 

---
## 文字設定
>   ```VBA
>   Range("A1").Font.Bold = true    '粗體字
> 
>   Range("B1").Font.Size = 20  '設定字體大小
> 
>   Range("C1").Interior.Color=RGB(0,255,0) '設定欄位顏色(顏色使用RGB表示)
> 
>   Range("D1").Font.Color = RGB(255, 0, 0) '設定字體顏色
> 
>   Range("E1").Borders.LineStyle = xlDouble    '外框設定成雙框線
> 
>   Range("F1").ColumnWidth = 30    '改變欄位寬度
> 
>   Range("G1").EntireColumn.AutoFit    '自動調整欄寬(需整欄選取 如果沒有資料則看不出變化)
> 
>   Range("H1").ClearContents   '清除資料內容
> 
>   Range("I1").ClearFormats    '清除資料格式
>   ```

---
## 複製、貼上
>   ```VBA
>   Range("A1:A2").Select   '選取要複製的範圍
>   Selection.Copy  '複製
>   Range("C3").Select  '選擇要貼上的位置
>   ActiveSheet.Paste   '貼上
>   ```
>   ---
>   ```VBA
>   Range("A1").Select  '選擇A1
>   Selection.Copy  '複製A1的內容
>   Range("B1","B4").Select '選擇B1~B4
>   Selection.PasteSpecial  '貼上
>   ```
>   **PasteSpecial 後面還能附加動作**
> 
>   ```
>   .PasteSpecial xlPasteFormats '貼上格式  
>   .PasteSpecial xlPasteValues '貼上值
>   .PasteSpecial Paste:=xlPasteValuesAndNumberFormats  '貼上值和數字格式
>   .PasteSpecial Paste:=xlPasteFormulas    '貼上公式
>   .PasteSpecial xlPasteFormulasAndNumberFormats   '貼上公式和數字格式
>   .PasteSpecial SkipBlanks:=True  '跳過空白  (只貼上有值的內容)
>   .PasteSpecial Transpose:=True   '轉置 (直的變橫的)
>   .PasteSpecial xlPasteValidation '資料驗證
>   .PasteSpecial xlPasteComments   '註解
>   ```  
>   其他請參考 [XlPasteType 列舉 (Excel)](https://docs.microsoft.com/zh-tw/office/vba/api/Excel.XlPasteType)


---
## 搜尋並取代
>   **在 A 工作表內，搜尋 B 工作表 R3 儲存格的值，並取代成 B 工作表 S3 儲存格的值**
>   ```VBA
>   Sheets("A").Cells.Replace _
>   Sheets("B").[R3], _
>   Sheets("B").[S3], xlPart   
>   ```

>   |  名稱 |  描述 |  
>   |  :-----: | :-----: |  
>   |   xlPart  |  與部分搜尋文字相符  |
>   |   xlWhole  |  與全部搜尋文字相符  |

---
## 自動填滿
>   ```AutoFill (Destination, 類型)```
>   ```VBA
>   Dim lrow As Long
>   lrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row
>   
>   Selection.AutoFill Destination:=Range("L2:L" & lrow)
>   ```

[XlAutoFillType 列舉 (Excel)](https://docs.microsoft.com/zh-tw/office/vba/api/excel.xlautofilltype)

---
## 自動篩選
>   ```AutoFilter (Field, Criteria1, Operator, Criteria2, SubField, VisibleDropDown) ```
>   ```VBA
>   ActiveSheet.Range("$A$1:$K$40").AutoFilter Field:=7, Criteria1:="1"
>   ```
>   
>   |  名稱  |  描述  |  
>   |  :-----:  | :-----:  |  
>   |   Field  |  要篩選的欄位  |
>   |   Criteria1  |  條件  |
>   |   Operator  |  指定篩選類型的 XlAutoFilterOperator 常數。  |
>   |   Criteria2  |  第二條件  |
>   |   SubField  |  第二條件  |
>   |   VisibleDropDown  |  是否顯示[自動篩選]下拉式箭號。|

* Criteria1：  
>   "="：找空白欄位。  
>   "<>"：找非空白欄位。  
>   "><"：找無資料欄位。  

* XlAutoFilterOperator：會指定用來關聯由篩選條件套用之兩個準則的運算子。

>   |  名稱 |  值  |  描述  |  
>   |  :-----:  |  :-----:  |  :-----:  |  
>   |  xlAnd |  1  |  Criteria1 與 Criteria2 的邏輯 AND  |  
>   |  xlBottom10Items |  4  |  在 Criteria1 中指定的專案數目，顯示最低值的專案  |  
>   |  xlBottom10Percent |  6  |  以 Criteria1) 指定的百分比顯示最低值的專案  |  
>   |  xlFilterCellColor |  8  |  儲存格的色彩  |  
>   |  xlFilterDynamic |  11  |  動態篩選條件  |  
>   |  xlFilterFontColor |  9  |  字型的色彩  |  
>   |  xlFilterIcon |  10  |  篩選條件圖示  |  
>   |  xlFilterValues |  7  |  篩選條件值  |  
>   |  xlOr |  第  |  Criteria1 或 Criteria2 的邏輯 OR  |  
>   |  xlTop10Items |  3  |  顯示最高值的專案 (在 Criteria1 中指定的專案數目)  |  
>   |  xlTop10Percent |  5  |  顯示最高值的專案 (以 Criteria1 指定的百分比)  |  
>   

---
## 排序
>   ```VBA
>   SortOn = xlSortOnValues   '要依什麼排序，預設為儲存格值
>   
>   Order = xlAscending  '升序
>   Order1:=xlDescending    '降序
>   
>   DataOption = xlSortNormal  
>   DataOption = xlSortTextAsNumbers    '將文字視為數字排序
>   
>   Header = xlYes   '有標題行（=1）
>   Header = xlGuess    '工作表自己判斷（=0）
>   MatchCase = False   '不區分大小寫
>   
>   Orientation = xlTopToBottom '按行排序
>   Orientation = xlLeftToRight '水平排序
>   
>   SortMethod = xlPinYin   '排序方法（使用拼音漢字排序）
>   SortMethod = xlStroke   '按每個字符的筆劃數排序
>   ```

---
## 錯誤處理（Error Handling）
>   * ```On Error Resume Next   ' 忽略錯誤，繼續執行```
>   * ```On Error GoTo ErrorHandler   ' 啟用錯誤處理機制```
>   
>   ```VBA
>   Sub Hello()
>     On Error GoTo ErrorHandler   ' 啟用錯誤處理機制
>     Dim x, y, z As Integer
>     x = 10
>     y = 0
>     z = x / y   ' 出現除以 0 的錯誤
>     MsgBox "z = " & z
>     Exit Sub    ' 結束子程序
>   
>   ErrorHandler:       ' 錯誤處理用的程式碼
>     MsgBox "錯誤 " & Err.Number & "：" & Err.Description
>     Resume Next       ' 繼續往下執行
>   
>   End Sub
>   ```
>

---
參考資料：  
* 基礎  
[VBA 程式設計](https://blog.gtwang.org/programming/vba/)  
[EXCEL VBA從頭來過-基本語法(上篇)](https://weilihmen.medium.com/excel-vba%E5%BE%9E%E9%A0%AD%E4%BE%86%E9%81%8E-%E5%9F%BA%E6%9C%AC%E8%AA%9E%E6%B3%95-%E4%B8%8A%E7%AF%87-c2bc76065ecd)  
[EXCEL VBA從頭來過-基本語法(中篇)](https://weilihmen.medium.com/excel-vba%E5%BE%9E%E9%A0%AD%E4%BE%86%E9%81%8E-%E5%9F%BA%E6%9C%AC%E8%AA%9E%E6%B3%95-%E4%B8%AD%E7%AF%87-4dda73e44eaf)  
[EXCEL VBA從頭來過-基本語法(下篇)](https://weilihmen.medium.com/excel-vba%E5%BE%9E%E9%A0%AD%E4%BE%86%E9%81%8E-%E5%9F%BA%E6%9C%AC%E8%AA%9E%E6%B3%95-%E4%B8%8B%E7%AF%87-cd3f6a389f34)  
[VBA + Excel VBA Code Examples](https://www.automateexcel.com/vba-code-examples/)  

* 列出所有文件  
[VBA code to loop through files in a folder (and sub folders)](https://exceloffthegrid.com/vba-code-loop-files-folder-sub-folders/)  
[Excel VBA 列出目錄中所有檔案、子目錄教學與範例](https://officeguide.cc/excel-vba-list-files-folders-in-directory-tutorial-examples/)  

* 匯出圖片  
[【VBA技巧】- N種方法從Excel中導出圖片，看這一篇就夠了](https://www.getit01.com/p2018013129070558/)  
[Excel VBA 7.78 Excel中的圖片如何批量保存？用VBA快如閃電](https://kknews.cc/zh-tw/career/plk9b2e.html)  

