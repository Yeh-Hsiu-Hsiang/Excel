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



參考資料：  
[INDEX 函數](https://support.microsoft.com/zh-tw/office/index-%E5%87%BD%E6%95%B8-a5dcf0dd-996d-40a4-a822-b56b061328bd#bmarray_form)  
[MATCH 函數](https://support.microsoft.com/zh-tw/office/match-%E5%87%BD%E6%95%B8-e8dffd45-c762-47d6-bf89-533f4a37673a)  

[Excel INDEX 函數用法教學：取出表格中特定位置的資料](https://blog.gtwang.org/windows/excel-index-function-tutorial-examples/)  
[Excel MATCH 函數用法教學：在表格中搜尋指定項目位置](https://blog.gtwang.org/windows/excel-match-function-tutorial/)  
