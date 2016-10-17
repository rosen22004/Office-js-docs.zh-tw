
# <a name="filetype-enumeration"></a>FileType 列舉
指定文件要傳回的格式。

|||
|:-----|:-----|
|**主應用程式︰**|PowerPoint、Word|
|**上次變更於**|1.1|

```js
Office.FileType
```


## <a name="members"></a>成員


**值**


|**列舉**|**值**|**描述**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|以 Office Open XML (OOXML) 格式傳回整個文件 (.pptx 或 .docx) ，成為位元組陣列。|
|Office.FileType.Pdf|"pdf"|以 PDF 格式傳回整個文件，成為位元組陣列。|
|Office.FileType.Text|"text"|只傳回文件的文字，成為**字串**。(只限 Word)|

## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 PowerPoint 和 Word 的支援。|
|1.1|新增儲存成 PDF 格式的支援。|
|1.0|已導入|
