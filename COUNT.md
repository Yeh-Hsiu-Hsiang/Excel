# COUNT 函數
## 計算儲存格中含有數字的儲存格個數
    COUNT(value1, [value2], ...)
* value1：必要。要計算的項目、儲存格參照或範圍。
* value2, ...：選擇性。最多 255 個。

Ex：    
>   =COUNT（A2：A7）  
>   →　計算儲存格 A2 到 A7 中含有數字的儲存格個數。

---
# COUNTA 函數
## 計算範圍中不是空白的儲存格個數
    COUNTA(value1, [value2], ...)
* value1：必要。要計算值的第一個引數。
* value2, ...：選擇性。要計算值的其他引數，最多有 255 個引數。

Ex：    
>   =COUNTA（A2：A6, B2：B6）  
>   →　計算儲存格 A2 到 A6 & B2 到 B6 中，非空白的儲存格數目。（A2~A6 + B2~B6 不為空白的儲存格數值）

---
# COUNTBLANK 函數
## 計算範圍中的空白儲存格個數  
    COUNTBLANK(Range)
* Range：必要。要計算空白儲存格的範圍。

Ex：    
>   =COUNTBLANK（A2：B4）  
>   →　計算 A2:B4 內的空白儲存格個數。

---
# COUNTIF 函數
## 計算符合條件的儲存格個數
    COUNTIF(range, criteria)
* range：必要。 要列入計算的儲存格。
* criteria：必要。 定義要將哪些儲存格列入計算的準則。

Ex：    
>   =COUNTIF（A2:A5,"apples"）  
>   →　計算儲存格 A2 到 A5 中有 apples 的儲存格個數。

---
# COUNTIFS 函數
## 將準則套用至多個範圍的儲存格，並計算所有準則均符合的個數
    COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2]…)
* criteria_range1：必要。 評估關聯準則的第一個範圍。
* criteria1：必要。 定義要將哪些儲存格列入計算的準則。
* criteria_range2， criteria2， ...：選用。 其他範圍及其相關準則。 最多允許 127 個範圍／準則組。

Ex：    
>   =COUNTIFS（B2:B5,"=是", C2:C5, "=否"）  
>   →　計算 B2:B5 儲存格資料為**是** __"且"__ C2:C5 儲存格資料為**否**的個數。

---
# 流水號
    TEXT(COUNTIF($C$2:C2, C2), "000")
>   COUNTIF（$C$2:C2, C2）：計算從儲存格 C2 以下的儲存格範圍，共有幾個和 C 欄各列相同的內容，這個數量即為其流水號。




---
參考資料：  
[COUNT 函數](https://support.microsoft.com/zh-tw/office/count-%E5%87%BD%E6%95%B8-a59cd7fc-b623-4d93-87a4-d23bf411294c)  
[COUNTA 函數](https://support.microsoft.com/zh-tw/office/counta-%E5%87%BD%E6%95%B8-7dc98875-d5c1-46f1-9a82-53f3219e2509)  
[COUNTBLANK 函數](https://support.microsoft.com/zh-tw/office/countblank-%E5%87%BD%E6%95%B8-6a92d772-675c-4bee-b346-24af6bd3ac22)   
[COUNTIFS 函數](https://support.microsoft.com/zh-tw/office/countifs-%E5%87%BD%E6%95%B8-dda3dc6e-f74e-4aee-88bc-aa8c2a866842)

[Excel 小教室 – 常用函數 IF、COUNTIF、COUNT、COUNTA 介紹](https://steachs.com/archives/29395)  
[Excel COUNT 系列函數懶人包，讓你統計數量更方便快速](https://today.line.me/tw/v2/article/qgpYDG)   
[Excel COUNTIF 與 COUNTIFS 函數用法教學：判斷多條件，計算數量](https://blog.gtwang.org/windows/excel-countif-countifs-function-tutorial/)  
[Excel-依項目自動排序號(COUNTIF)](https://isvincent.pixnet.net/blog/post/39030187-excel-%E4%BE%9D%E9%A0%85%E7%9B%AE%E8%87%AA%E5%8B%95%E6%8E%92%E5%BA%8F%E8%99%9F(countif))
 



 
