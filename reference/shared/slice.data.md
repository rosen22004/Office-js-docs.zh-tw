
# <a name="slice.data-property"></a>Slice.data 屬性
取得檔案配量的未經處理資料。

|||
|:-----|:-----|
|**主應用程式︰**|PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|檔案|
|**上次變更於**|1.1|

```
var sliceData = slice.data;
```


## <a name="return-value"></a>傳回值

已透過呼叫 **Document.getFileAsync** 方法的 **fileType** 參數，指定_Office.FileType.Text_ ("text") 或 [Office.FileType.Compressed](../../reference/shared/document.getfileasync.md) ("compressed") 格式的檔案配量未經處理資料。


## <a name="remarks"></a>備註

“compressed” 格式的檔案會回傳位元組陣列，其可視需要轉換為以 base64 編碼的字串。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

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



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 PowerPoint 和 Word 的支援。|
|1.0|已導入|
