
# <a name="document.getactiveviewasync-method"></a>Document.getActiveViewAsync 方法
 傳回簡報的目前檢視狀態 (編輯或讀取)。

|||
|:-----|:-----|
|**主應用程式︰**Excel、PowerPoint、Word|**增益集類型：**內容、工作窗格|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|ActiveView|
|**已新增於 ActiveView**|1.1|

```
Office.context.document.getActiveViewAsync([,options], callback);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **getActiveViewAsync** 方法的回呼函數中，[AsyncResult.value](../../reference/shared/asyncresult.value.md) 屬性會傳回簡報的目前檢視狀態。傳回的值可能是 `edit` 或 `read`。`edit` 會對應至您可編輯投影片的任何一個檢視，例如**一般**或**大綱模式**。`read` 會對應至**投影片放映**或**閱讀檢視**。


## <a name="remarks"></a>備註

可在檢視變更時，觸發事件。


## <a name="example"></a>範例

若要檢視目前簡報，您需要寫入傳回該值的回呼函數。下列範例顯示如何：


-  **傳遞匿名的回呼函數**，可將檢視類型傳回至 _getActiveViewAsync_ 方法的 **callback** 參數。
    
-  在增益集頁面上**顯示值**。
    

```js
function getFileView() {
    // Get whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage(asyncResult.value);
        }
    });
}
```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|||Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|||Y|

|||
|:-----|:-----|
|**可用於需求集合**|ActiveView|
|**已新增於 ActiveView**|1.1|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄





****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|已導入。|
