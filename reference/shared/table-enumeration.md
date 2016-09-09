
# Table 列舉
指定_資料表格式化方法_的 [cellFormat](../../docs/excel/format-tables-in-add-ins-for-excel.md) 參數中的 `cells:` 屬性列舉值。

|||
|:-----|:-----|
|**主機︰**|Excel|
|**已新增**|1.1|

```
Office.Table
```

## 成員


**值**


|**列舉**|**值**|**說明**|
|:-----|:-----|:-----|
|Office.Table.All|"all"|整張表格，若有欄標題、資料和總計，也包含在內。|
|Office.Table.Data|"data"|只有資料 (沒有標題和總計)。|
|Office.Table.Headers|"headers"|只有標題列。|

## 支援詳細資料


下列矩陣中的大寫 Y，表示已在相對應的 Office 主應用程式中支援列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 的支援。|
|1.1|已導入|
