
# ProjectProjectFields 列舉
指定可作為 **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)** 方法參數的專案欄位。

|||
|:-----|:-----|
|**主機︰**|Project|
|**已新增於**|1.0|

```
ProjectProjectFields={
    CurrencyDigits: 0, 
    CurrencySymbol: 1, 
    CurrencySymbolPosition: 2, 
    DurationUnits: 3,
    GUID: 4, 
    Finish: 5, 
    Start: 6, 
    ReadOnly: 7, 
    VERSION: 8, 
    WorkUnits: 9, 
    ProjectServerUrl: 10, 
    WSSUrl: 11, 
    WSSList: 12
}
```


## 成員


****


|**成員	**|**說明**|
|:-----|:-----|
|**CurrencyDigits**|貨幣小數點後的位數。|
|**CurrencySymbol**|貨幣符號。|
|**CurrencySymbolPosition**|貨幣符號的位置︰未指定 =-1；值之前沒有空格 ($0) = 0；值之後沒有空格 (0$) = 1；值之前有一個空格 ($ 0) = 2；值之後有一個空格 (0 $) = 3。|
|**GUID**|專案的 GUID。|
|**Finish**|專案完成日期。|
|**Start**|專案開始日期。|
|**ReadOnly**|指定專案是否為唯讀。|
|**VERSION**|專案版本。|
|**WorkUnits**|專案的工時單位，如天或小時。|
|**ProjectServerUrl**|儲存於 Project Server 之專案的 Project Web Appy URL。|
|**WSSUrl**|與 SharePoint 清單同步之專案的 SharePoint URL。|
|**WSSList**|與任務清單同步之專案的 SharePoint 清單名稱。|

## 備註

**ProjectProjectFields** 常數可作為 **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)** 方法的參數。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|

## 請參閱



#### 其他資源


[getProjectFieldAsync 方法](../../reference/shared/projectdocument.getprojectfieldasync.md)
