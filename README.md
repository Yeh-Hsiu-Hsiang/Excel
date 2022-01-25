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
    >>  檢查儲存格 A6 中的值 - #N/A 是否為 #N/A 錯誤。

## COUNTIF
>   COUNTIF(range, criteria)

## FIND
>   FIND(find_text, within_text, [start_num])

## SUBSTITUTE
>   SUBSTITUTE(text, old_text, new_text, [instance_num])

## ADDRESS
>   ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])
    >>  A1：選擇性。指定 A1 或 R1C1 欄名列號表示法的邏輯值。  
    >   
    >>  sheet_text：選擇性。文字值，指定要用作外部參考之工作表的名稱。

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




---
參考資料：
[IS 函數](https://support.microsoft.com/zh-tw/office/is-%E5%87%BD%E6%95%B8-0f2d7971-6019-40a0-a171-f2d869135665)  
[Excel REPLACE 與 SUBSTITUTE 函數用法教學：字串取代，自動修改文字資料](https://blog.gtwang.org/windows/excel-replace-substitute-function-tutorial/)  
[ADDRESS 函數](https://support.microsoft.com/zh-tw/office/address-%E5%87%BD%E6%95%B8-d0c26c0d-3991-446b-8de4-ab46431d4f89)  
