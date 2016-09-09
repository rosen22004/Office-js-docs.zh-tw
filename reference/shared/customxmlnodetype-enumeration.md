
# CustomXMLNodeType 列舉
指定節點類型。



|||
|:-----|:-----|
|**主機︰**|Word|
|**上次變更於**|1.1|



```js
Office.CustomXMLNodeType
```


## 成員


**值**


|**列舉**|**值**|**說明**|
|:-----|:-----|:-----|
|Office.CustomXMLNodeType.Attribute|"attribute"|節點為屬性。|
|Office.CustomXMLNodeType.CData|"CData"|節點為 CData 類型。|
|Office.CustomXMLNodeType.NodeComment|"comment"|節點為註解。|
|Office.CustomXMLNodeType.Element|"element"|節點為元件。|
|Office.CustomXMLNodeType.NodeDocument|"nodeDocument"|節點為文件元素。|
|Office.CustomXMLNodeType.ProcessingInstruction|"processingInstruction"|節點為處理指示。|
|Office.CustomXMLNodeType.Text|"text"|節點為文字節點。|

## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|



|||
|:-----|:-----|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Word 的支援。|
|1.0|已導入|
