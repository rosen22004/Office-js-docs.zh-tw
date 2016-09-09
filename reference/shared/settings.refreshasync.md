

# Settings.refreshAsync 方法
讀取保存在文件內的所有設定，並重新整理保留在記憶體內部的內容或工作窗格增益集複本。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**上次變更於**|1.1|

```js
Office.context.document.settings.refreshAsync(callback);
```


## 參數

_回呼_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**物件**

&nbsp;&nbsp;&nbsp;&nbsp;回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。

    



## 回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **refreshAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用途**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取具重新整理過的值的 [Settings](../../reference/shared/settings.md) 物件。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## 備註

相同增益集的多個執行個體處理相同文件時，此方法便可用於 Word 與 PowerPoint 共同編寫情節中。因各增益集在使用者開啟文件時，會使用文件載入的設定記憶體內部複本，所以可同步各使用者使用的設定值。每當增益集的執行個體呼叫 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法，將所有使用者的設定保存至文件時，就可能發生此狀況。從增益集的 **settingsChanged** 事件的事件處理常式，呼叫 [refreshAsync](../../reference/shared/settings.settingschangedevent.md) 方法，會重新整理所有使用者的設定值。

可從為 Excel 而建立的增益集中呼叫 **refreshAsync** 方法，但因該方法不支援共同編寫，所以沒有必要呼叫該方法。


## 範例




```js
function refreshSettings() {
    Office.context.document.settings.refreshAsync(function (asyncResult) {
        write('Settings refreshed with status: ' + asyncResult.status);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|Settings|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增支援在 Access 內容增益集中自訂設定。|
|1.0|已導入|
