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
## 註解：```'```
```VBA
'MsgBox ("Hello, world!")
```

---
## 換行：```_```
```VBA
x = 1 + 2 + 3 + _
    4 + 5 + 6
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
>
>   WorkSheets  '所有工作表
>
>   Worksheets.Add  '新增工作表
>
>   Worksheets(1).Name = "新的工作表"   '改第一個工作表名稱
>
>   WorkSheets(1).Activate  '第一個工作表拉到當前視窗
>   
>   Ex:
>   Worksheets("活頁簿1").Activate
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
>   Header = xlYes   '有標題行
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
[VBA 程式設計](https://blog.gtwang.org/programming/vba/)  
[EXCEL VBA從頭來過-基本語法(上篇)](https://weilihmen.medium.com/excel-vba%E5%BE%9E%E9%A0%AD%E4%BE%86%E9%81%8E-%E5%9F%BA%E6%9C%AC%E8%AA%9E%E6%B3%95-%E4%B8%8A%E7%AF%87-c2bc76065ecd)  
[EXCEL VBA從頭來過-基本語法(中篇)](https://weilihmen.medium.com/excel-vba%E5%BE%9E%E9%A0%AD%E4%BE%86%E9%81%8E-%E5%9F%BA%E6%9C%AC%E8%AA%9E%E6%B3%95-%E4%B8%AD%E7%AF%87-4dda73e44eaf)  
[EXCEL VBA從頭來過-基本語法(下篇)](https://weilihmen.medium.com/excel-vba%E5%BE%9E%E9%A0%AD%E4%BE%86%E9%81%8E-%E5%9F%BA%E6%9C%AC%E8%AA%9E%E6%B3%95-%E4%B8%8B%E7%AF%87-cd3f6a389f34)  
[VBA + Excel VBA Code Examples](https://www.automateexcel.com/vba-code-examples/)  
[Excel VBA 除錯技巧：Debug.Print 與即時運算視窗使用教學](https://officeguide.cc/excel-vba-debug-immediate-window-tutorial/)