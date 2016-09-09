
# Settings.saveAsync 方法
將設定屬性包的記憶體內部複本保存於文件中。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**上次變更於**|1.1|

```js
Office.context.document.settings.saveAsync(callback);
```


## 參數



_回呼_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**物件**

&nbsp;&nbsp;&nbsp;&nbsp;回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。 選用。

    



## 回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **saveAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用途**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|因為沒有可擷取的物件或資料，所以一律傳回 **undefined**。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## 備註

初始化增益集時會載入增益集先前儲存的所有設定，所以在工作階段的存留期間內，您只要使用 [set](../../reference/shared/settings.set.md) 和 [get](../../reference/shared/settings.get.md) 方法，就能使用設定屬性包的記憶體內部複本。如果您想要保存設定，以便在下次使用增益集時使用，請使用 **saveAsync** 方法。


 >**附註**：**saveAsync** 方法會將記憶體內部設定屬性包保存至文件檔案中，但僅會在使用者 (或 **AutoRecover** 設定) 將文件儲存至檔案系統時，儲存文件檔案本身的變更。

相同增益集的其他執行個體可能變更該設定，且所有執行個體應皆可使用這些變更時，[RefreshAsync](../../reference/shared/settings.refreshasync.md) 方法便僅可用於共同編寫情節 (僅在 Word 中支援)。


## 範例




```js
function persistSettings() {
    Office.context.document.settings.saveAsync(function (asyncResult) {
        write('Settings saved with status: ' + asyncResult.status);
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
