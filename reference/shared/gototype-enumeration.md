
# GoToType 列舉
指定要瀏覽至的位置或物件類型。

|||
|:-----|:-----|
|**主機︰**|Excel、PowerPoint、Word|
|**已新增於**|1.1|

```js
Office.GoToType
```


## 成員


**值**


|**列舉**|**值**|**說明**|**支援的用戶端**|
|:-----|:-----|:-----|:-----|
|Office.GoToType.Binding|"binding"|移至使用指定繫結識別碼的繫結物件。|Excel</br>Word|
|Office.GoToType.NamedItem|"namedItem"|移至使用該項目名稱的項目，例如指派給資料表或範圍的名稱。在 Excel 中，您可以針對已命名的範圍或資料表，使用任何結構化的參考︰"Worksheet2!Table1"|Excel|
|Office.GoToType.Slide|"slide"|移至使用指定識別碼的投影片。|PowerPoint|
|Office.GoToType.Index|"index"|依據投影片號碼或列舉，移至指定的索引：</br>**Office.Index.First**</br>**Office.Index.Last**</br>**Office.Index.Next**</br>**Office.Index.Previous**|PowerPoint|

## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。


如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|已導入|
