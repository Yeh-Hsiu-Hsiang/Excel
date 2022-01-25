# ROUND 函數
## 四捨五入
    ROUND(number, num_digits)
* number：必要。 要四捨五入的數字。
* num_digits：必要。 要進位的位數。

Ex：    
>   =ROUND(12345.6789, 2)  
>   →　四捨五入到小數第二位。（=12345.68）

---
# ROUNDUP 函數
## 無條件進位
    ROUNDUP(number, num_digits)
* number：必要。 要無條件進位的數字。
* num_digits：必要。 要進位的位數。

Ex：    
>   =ROUNDUP(12345.6789,0)  
>   →　無條件進位到整數。（=12346）

---
# ROUNDDOWN 函數
## 無條件捨去  
    ROUNDDOWN(number, num_digits)
* number：必要。 要無條件捨去的數字。
* num_digits：必要。 要進位的位數。

Ex：    
>   =ROUNDDOWN(12345.6789,-2)  
>   →　無條件捨去到百位數。（=12300）


---
|  num_digits |  進位的位數 |  
|  :-----: | :-----: |  
|   2  |  小數第二位  |
|   1  |  小數第一位  |
|   0  |  整數  |
|   -1  |  十位數  |
|   -2  |  百位數  |


---
參考資料：  
[ROUND 函數](https://support.microsoft.com/zh-tw/office/round-%E5%87%BD%E6%95%B8-c018c5d8-40fb-4053-90b1-b3e7f61a213c)  
[ROUNDUP 函數](https://support.microsoft.com/zh-tw/office/roundup-%E5%87%BD%E6%95%B8-f8bc9b23-e795-47db-8703-db171d0c42a7)  
[ROUNDDOWN 函數](https://support.microsoft.com/zh-tw/office/rounddown-%E5%87%BD%E6%95%B8-2ec94c73-241f-4b01-8c6f-17e6d7968f53) 
 
