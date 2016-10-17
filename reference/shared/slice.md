
# <a name="slice-object"></a>Slice 物件
代表文件檔案的配量。

|||
|:-----|:-----|
|**主應用程式︰**|PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|檔案|
|**上次變更於**|1.1|

```
slice
```


## <a name="members"></a>成員


**屬性**


|**名稱**|**描述**|
|:-----|:-----|
|**[data](../../reference/shared/slice.data.md)**|取得檔案配量的未經處理資料。|
|**[index](../../reference/shared/slice.index.md)**|取得檔案配量的索引。|
|**[size](../../reference/shared/slice.size.md)**|取得配量大小，以位元組為單位。|

## <a name="remarks"></a>備註

可透過 **File.getSliceAsync** 方法存取 [Slice](../../reference/shared/file.getsliceasync.md) 物件。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|


|||
|:-----|:-----|
|**可用於需求集合**|檔案|
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 PowerPoint 和 Word 的支援。|
|1.0|已導入|
