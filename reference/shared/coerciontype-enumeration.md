
# CoercionType 列舉
指定如何強制轉型所傳回或由叫用方法設定的資料。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於信箱**|1.1|

```js
Office.CoercionType
```

## 成員


**值**


|**列舉**|**值**|**說明**|
|:-----|:-----|:-----|
|Office.CoercionType.Html|"html"|傳回或將資料設為 HTML。<br/><br/> **附註：**僅適用於 Word 增益集和 Outlook 之 Outlook 增益集中的資料 (撰寫模式)。|
|Office.CoercionType.Matrix|"matrix"|傳回或將資料設為沒有標頭的表格式資料。 資料已傳回或設為包含一連串一維字元之陣列的陣列。 例如，兩欄中的三列**字串**值將會是：` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`。<br/><br/> **附註：**僅適用於 Excel 和 Word 中的資料。|
|Office.CoercionType.Ooxml|"ooxml"|傳回或將資料設為 Office Open XML。<br/><br/> **附註：**僅適用於 Word 中的資料。|
|Office.CoercionType.SlideRange|"slideRange"|傳回包含所選取投影片之 id、標題和索引之陣列的 JSON 物件。例如：`{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` 適用於選取兩張投影片。<br/><br/> **附註：**僅適用於呼叫 [Document.getSelectedData](../../reference/shared/document.getselecteddataasync.md) 方法以取得目前投影片或選取的投影片範圍時，PowerPoint 中的資料。|
|Office.CoercionType.Table|"table"|傳回或將資料設為帶有選擇性標頭的表格式資料。 資料已傳回或設為帶有選擇性標頭之陣列的陣列。<br/><br/> **附註：**僅適用於 Access、Excel 和 Word 中的資料。|
|Office.CoercionType.Text|"text"|傳回或將資料設為文字 (**字串**)。資料已傳回或設為一連串一維字元。|
|Office.CoercionType.Image|"image"|資料已傳回或設為影像資料流。<br/><br/> **附註：**僅適用於 Excel、Word 和 PowerPoint 中的資料。|
PowerPoint 僅支援 **Office.CoercionType.Text**、**Office.CoercionType.Image** 和 **Office.CoercionType.SlideRange**。

Project 僅支援 **Office.CoercionType.Text**。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Office for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|是|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**增益集類型**|內容、Outlook (撰寫模式)、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 Word Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增對 Access 增益集的支援。|
|1.1|新增對[撰寫模式 Outlook 增益集](../../docs/outlook/compose-scenario.md)的支援。|
|1.0|已導入|
