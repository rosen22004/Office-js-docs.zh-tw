
# ValueFormat 列舉
指定叫用方法所傳回的值，例如數字和日期，是否以其套用的格式設定傳回。

|||
|:-----|:-----|
|**主機︰**|Excel、Project、Word|
|**已新增於**|1.0|

```
Office.ValueFormat
```


## 成員


**值**


|**列舉**|**值**|**說明**|
|:-----|:-----|:-----|
|Office.ValueFormat.Formatted|"formatted"|傳回格式化的資料。|
|Office.ValueFormat.Unformatted|"unformatted"|傳回未格式化的資料。|

## 備註

例如，將 _valueFormat_ 參數指定為 `"formatted"`，主應用程式中格式化成貨幣的數字，或格式化為 mm/dd/yy 的日期，將保留其格式設定。如果將 _valueFormat_ 參數指定為 `"unformatted"`，會以基本序列序號傳回日期。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.0|已導入|
