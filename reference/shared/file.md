
# <a name="file-object"></a>File 物件
代表與 Office 增益集相關聯的文件檔案。

|||
|:-----|:-----|
|**主應用程式︰**|PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|檔案|
|**上次變更於**|1.1|

```
file
```


## <a name="members"></a>成員


**屬性**


|**名稱**|**描述**|
|:-----|:-----|
|**[size](../../reference/shared/file.size.md)**|取得文件的檔案大小 (位元組)。|
|**[sliceCount](../../reference/shared/file.slicecount.md)**|取得檔案被劃分的配量數。|

**方法**


|**名稱**|**描述**|
|:-----|:-----|
|**[closeAsync](../../reference/shared/file.closeasync.md)**|關閉文件檔案。|
|**[getSliceAsync](../../reference/shared/file.getsliceasync.md)**|傳回指定的配量。|

## <a name="remarks"></a>備註

在傳遞至 **Document.getFileAsync** 方法的回呼函數中，使用 [AsyncResult.value](../../reference/shared/asyncresult.value.md) 屬性以存取 [File](../../reference/shared/document.getfileasync.md) 物件。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||||
|:-----|:-----|:-----|:-----|
||Office for Windows desktop|Office Online (在瀏覽器中)|Office for iPad|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**可用於需求集合**|檔案|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 PowerPoint 和 Word 的支援。|
|1.0|已導入|
