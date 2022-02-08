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
>   >   |  abs_num |  傳回此參照類型 |  
>   >   |  :-----: | :-----: |  
>   >   |   1 或省略  |  絕對參照  |
>   >   |   2  |  絕對列；相對欄  |
>   >   |   3  |  相對列；絕對欄  |
>   >   |   4  |  相對參照  |
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
# LOOKUP 函數
## 在單列或單欄範圍 (亦稱為向量) 中尋找值，並傳回第二個單列或單欄範圍內相同位置的值。
    LOOKUP(lookup_value, lookup_vector, [result_vector])
* lookup_value：必要。 要查閱的值。
* lookup_vector：必要。 單列或單欄的範圍，在 lookup_vector 中的數值必須以遞增次序排列。
* result_vector：選填。　單列或單欄的範圍。大小應與 lookup_vector 相同。

Ex：    
>   =LOOKUP(E13, A13:A20,B13:B20)  
>   →　在 A13:A20 範圍中，尋找 E13 的資料，並傳回B欄對應的值。

---
# VLOOKUP 函數
## 在陣列或表格的最左欄中尋找含有某特定值的欄位，再傳回同一列中某一指定儲存格中的值。
    VLOOKUP (lookup_value, table_array, col_index_num, [range_lookup])
* lookup_value：必要。 要查閱的值。
* table_array：必要。 搜尋及傳回值的儲存格範圍。
* col_index_num：必要。　數值，代表所要傳回的值位於 table_array 中的第幾欄。
* range_lookup：選填。 指定要尋找大約符合或完全符合值的邏輯值。

>   |  range_lookup |  傳回此參照類型 |  
>   |  :-----: | :-----: |  
>   |   0 / FALSE  |  完全相符，傳回完全符合的值  |
>   |   1 / TRUE  |  大約相符，傳回部分符合的值  |

Ex：    
>   = VLOOKUP("連",B2:E7,2,FALSE)  
>   →　在 B2:E7 範圍中，尋找完全符合"連"的資料，並傳回第二欄的值。


---
參考資料：  
[INDEX 函數](https://support.microsoft.com/zh-tw/office/index-%E5%87%BD%E6%95%B8-a5dcf0dd-996d-40a4-a822-b56b061328bd#bmarray_form)  
[MATCH 函數](https://support.microsoft.com/zh-tw/office/match-%E5%87%BD%E6%95%B8-e8dffd45-c762-47d6-bf89-533f4a37673a)  
[ADDRESS 函數](https://support.microsoft.com/zh-tw/office/address-%E5%87%BD%E6%95%B8-d0c26c0d-3991-446b-8de4-ab46431d4f89)  
[LOOKUP 函數](https://support.microsoft.com/zh-tw/office/lookup-%E5%87%BD%E6%95%B8-446d94af-663b-451d-8251-369d5e3864cb)  
[VLOOKUP 函數](https://support.microsoft.com/zh-tw/office/vlookup-%E5%87%BD%E6%95%B8-0bbc8083-26fe-4963-8ab8-93a18ad188a1) 

[Excel INDEX 函數用法教學：取出表格中特定位置的資料](https://blog.gtwang.org/windows/excel-index-function-tutorial-examples/)  
[Excel MATCH 函數用法教學：在表格中搜尋指定項目位置](https://blog.gtwang.org/windows/excel-match-function-tutorial/)  
[Excel Address函數用法的7個實例，含四種引用類型](https://kknews.cc/zh-tw/career/xgjejjr.html)  
[Excel VLOOKUP 函數教學：按列搜尋表格，自動填入資料](https://blog.gtwang.org/windows/excel-vlookup-function-tutorial/)  
[有重複值又如何？VLOOKUP同樣輕鬆查找！](https://kknews.cc/career/vvbl5ll.html)  

 

