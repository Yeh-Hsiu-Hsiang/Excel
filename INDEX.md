# INDEX 函數

>   ### VLOOKUP 函數只能用於查詢欄在最左邊使用。若要從中間欄查詢資料必須用 INDEX 函數。

## INDEX：傳回表格或範圍內的值。  
    INDEX（array, row_num, [column_num]） 
    Index（範圍, 傳回值的列, 傳回值的欄）
 
 
| 商品 |  價格 |  數量  | 
| :-----:  | :-----: |  :-----: |  
| USB  |  300  |  10  |    
| 手機  |  30000  |  5 |    
| MP3  |  15000  |  15 |    

Ex：
>   =INDEX（A2:C4, 2, 2）  
>   →　在 A2:C4 範圍內，第二列及第二欄交叉點的值。（=30000）

### **INDEX 函數常與 MATCH 函數一起使用。**

---
# MATCH 函數
    MATCH（lookup_value, lookup_array, [match_type]）
    MATCH（搜尋值, 搜尋範圍, 搜尋類型）

>    搜尋類型：  
>    若為 0 則代表尋找完全一樣的值。  
>    若為 1 則代表尋找小於或等於搜尋值的最大值。  
>    若為 -1 則代表尋找大於或等於搜尋值的最小值。  

Ex：
>   MATCH（30000, $B:$B, 0, 1）  
>   →　在 B 欄找尋 30000 的位置。（=3）


>   =INDEX（$A:$A, MATCH（30000,$B:$B, 0）, 1）  
>   →　在 B 欄找尋 30000 的位置後，再從 A 欄裡面找第 3 列的值。 （= 手機）  

---
# ADDRESS 函數
    ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])

>  abs_num：選擇性。 指定要傳回之參照類型的數值。
    >>   |  abs_num |  傳回此參照類型 |  
    >>   |  :-----: | :-----: |  
    >>   |   1 或省略  |  絕對參照  |
    >>   |   2  |  絕對列；相對欄  |
    >>   |   3  |  相對列；絕對欄  |
    >>   |   4  |  相對參照  |
>
>  A1：選擇性。指定 A1 或 R1C1 欄名列號表示法的邏輯值。  
>   
>  sheet_text：選擇性。文字值，指定要用作外部參考 之工作表的名稱。

Ex：
>   ADDRESS（2, 3, 2）  
>   →　傳回第二列絕對列第三欄相對欄。（=C$2）

* Indirect + Address + Row + Column：提取由指定行開始的數據。
>   =INDIRECT(ADDRESS(ROW($A2), COLUMN()))  
![提取由指定行開始的數據](http://i2.kknews.cc/F-LQe35DtNGL7tJwL3evikBBpDRdQi2ywQ/0.jpg "Indirect + Address + Row + Column")

* OffSet + Indirect + Address + Match：查找資料。
>   =OFFSET(INDIRECT(ADDRESS(MATCH(A10,A1:A6,0),1)),,1)  
>   ![查找資料](http://i1.kknews.cc/KmMbbRzoYI1b75KGLKcVb3AImgZApnJosA/0.jpg "OffSet + Indirect + Address + Match")  

* Sum + OffSet + Indirect + Address：多表格求總和。
>   =SUM(OFFSET(INDIRECT(ADDRESS(1,4,,,A2&"月")),1,,6))  
>   ![多表格求總和](http://i2.kknews.cc/1f7PRkzO9HYhSgHOXNXuan7tUDc6Jk9JHw/0.jpg "Sum + OffSet + Indirect + Address")  
>   返回高度為 6、寬度為 1 的儲存格引用。



---
參考資料：  
[INDEX 函數](https://support.microsoft.com/zh-tw/office/index-%E5%87%BD%E6%95%B8-a5dcf0dd-996d-40a4-a822-b56b061328bd#bmarray_form)  
[MATCH 函數](https://support.microsoft.com/zh-tw/office/match-%E5%87%BD%E6%95%B8-e8dffd45-c762-47d6-bf89-533f4a37673a)  
[Excel INDEX 函數用法教學：取出表格中特定位置的資料](https://blog.gtwang.org/windows/excel-index-function-tutorial-examples/)  
[Excel MATCH 函數用法教學：在表格中搜尋指定項目位置](https://blog.gtwang.org/windows/excel-match-function-tutorial/)  
[Excel Address函數用法的7個實例，含四種引用類型](https://kknews.cc/zh-tw/career/xgjejjr.html)  

