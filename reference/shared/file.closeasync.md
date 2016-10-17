
# <a name="file.closeasync-method"></a>File.closeAsync 方法
關閉文件檔案。

|||
|:-----|:-----|
|**主應用程式︰**|PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|檔案|
|**已新增於**|1.1|

```js
File.closeAsync(callback);
```


## <a name="parameters"></a>參數


_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**物件**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;回呼傳回時所叫用的函數，其唯一的參數為 [AsyncResult](../../reference/shared/asyncresult.md) 類型。選用。
    

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **closeAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|因為沒有可擷取的物件或資料，所以一律傳回 **undefined**。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

不允許記憶體中有超過兩個文件；否則 [Document.getFileAsync](../../reference/shared/document.getfileasync.md) 作業會失敗。當您完成使用檔案時，請使用 **File.closeAsync**方法關閉檔案。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|檔案|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 PowerPoint 和 Word 的支援。|
|1.0|已導入|
