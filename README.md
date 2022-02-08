# 常用函數

## ROW 函數
>   ROW([reference])

## COLUMN 函數
>   COLUMN([reference])

## INDEX 函數  
>   INDEX（array, row_num, [column_num]） 

## MATCH 函數
>   MATCH（lookup_value, lookup_array, [match_type]）

## IFERROR
>   IFERROR(value, value_if_error)

## ISNA
>   =ISNA(A6)
>   >  檢查儲存格 A6 中的值 - #N/A 是否為 #N/A 錯誤。

## COUNTIF
>   COUNTIF(range, criteria)

## COUNTIFS
>   COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2]…)
>   
>   ## 流水號
>   ```TEXT(COUNTIF($C$2:C2, C2), "000")```  
>   →　計算從儲存格 C2 以下的儲存格範圍，共有幾個和 C 欄各列相同的內容，這個數量即為其流水號。

# VLOOKUP 
>   VLOOKUP (lookup_value, table_array, col_index_num, [range_lookup])

## FIND
>   FIND(find_text, within_text, [start_num])  
>    
>   >   ## Ex: 2022/1/20_AZ_1_2_3 要分成施打日、施打疫苗種類、第幾劑  
>   >   施打日：```=LEFT($G3, FIND("_",$G3)-1)```  
>   >   疫苗種類：```=MID($G3, FIND("_",$G3)+1, (FIND( "_", $G3, FIND("_", $G3)+1))-(FIND("_",$G3)+1))```  
>   >   第幾劑：```=MID($G3, FIND("_", $G3, FIND("_", $G3)+1)+1, 5)```  
>   >   →　FIND("_",$G3)：找到第一個 _ 的位置。  
>   >   →　FIND("_", $G3)+1：找到第二個 _ 的位置。  

## SUBSTITUTE
>   SUBSTITUTE(text, old_text, new_text, [instance_num])

## ADDRESS
>   ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])  
>    >  abs_num：選擇性。 指定要傳回之參照類型的數值。 
>    >   
>   |  abs_num |  傳回此參照類型 |  
>   |  :-----: | :-----: |  
>   |   1 或省略  |  絕對參照  |  
>   |   2  |  相對列；絕對欄  |  
>   |   3  |  絕對列；相對欄  |  
>   |   4  |  相對參照  |  
>    >
>    >  A1：選擇性。指定 A1 或 R1C1 欄名列號表示法的邏輯值。  
>    >   
>    >  sheet_text：選擇性。文字值，指定要用作外部參考之工作表的名稱。

## OFFSET
>   OFFSET(reference, rows, cols, [height], [width])

## SUMPRODUCT
>   =SUMPRODUCT(array1, [array2], [array3], ...)

## ROUND
>   ROUND(number, num_digits)
>   * ROUNDUP
>        >    ROUNDUP(number, num_digits)
>   * ROUNDDOWN
>        >    ROUNDDOWN(number, num_digits)
>
>   |  num_digits |  進位的位數 |  
>   |  :-----: | :-----: |  
>   |   2  |  小數第二位  |
>   |   1  |  小數第一位  |
>   |   0  |  整數  |
>   |   -1  |  十位數  |
>   |   -2  |  百位數  |

## SUBTOTAL
>   SUBTOTAL(function_num,ref1,[ref2],...)
>    >  Function_num：必要。數字 1-11 或 101-111 指定要用於計算小計的函數。  
>   
>   |  Function_num (包含隱藏的列)  |  Function_num (忽略隱藏列) |  函數  |
>   |  :-----: | :-----: |  :-----: |  
>   |   1  |  101  |    AVERAGE     |
>   |   2  |  102  |    COUNT       |
>   |   3  |  103  |    COUNTA      |
>   |   4  |  104  |    MAX     |
>   |   5  |  105  |    MIN     |
>   |   6  |  106  |    PRODUCT     |
>   |   7  |  107  |    STDEV       |
>   |   8  |  108  |    STDEVP      |
>   |   9  |  109  |    SUM     |
>   |   10  |  110  |   VAR     |
>   |   11  |  111  |   VARP        |



---
參考資料：  
[IS 函數](https://support.microsoft.com/zh-tw/office/is-%E5%87%BD%E6%95%B8-0f2d7971-6019-40a0-a171-f2d869135665)  
[Excel REPLACE 與 SUBSTITUTE 函數用法教學：字串取代，自動修改文字資料](https://blog.gtwang.org/windows/excel-replace-substitute-function-tutorial/)    
[聯成電腦分享：Excel擷取所需字元、字串（下）](https://www.lccnet.com.tw/lccnet/article/details/1958)  
[SUBTOTAL 函數](https://support.microsoft.com/zh-tw/office/subtotal-%E5%87%BD%E6%95%B8-7b027003-f060-4ade-9040-e478765b9939)  
[有重複值又如何？VLOOKUP同樣輕鬆查找！](https://kknews.cc/career/vvbl5ll.html)
