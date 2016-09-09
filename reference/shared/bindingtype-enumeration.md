
# BindingType 列舉
 指定應傳回的繫結物件類型。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Word|
|**上次變更**|1.1|

```
Office.BindingType
```


## 成員


**值**


|**列舉**|**值**|**說明**|
|:-----|:-----|:-----|
|Office.BindingType.Matrix|"matrix"|沒有標題列的表格式資料。資料會以陣列的陣列傳回，例如以此形式：` [[row1column1, row1column2],[row2column1, row2column2]]`|
|Office.BindingType.Table|"table"|帶有標題列的表格式資料。資料會以 [TableData](../../reference/shared/tabledata.md) 物件傳回。|
|Office.BindingType.Text|"text"|純文字。資料會以一連串的字元傳回。|

## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|Y|||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.1|新增支援在 Access 增益集中繫結表格資料。|
|1.0|已導入。|
