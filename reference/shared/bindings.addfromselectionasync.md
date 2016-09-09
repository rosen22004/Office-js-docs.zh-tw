
# Bindings.addFromSelectionAsync 方法
新增繫結至文件中目前的選取範圍。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**上次變更**|1.1|

```
bindingsObj.addFromSelectionAsync(bindingType [, options], callback);
```


## 參數



|**名稱**|**類型	**|**說明**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|指定要建立的繫結物件類型。必要。如果選取的物件無法強制轉型至指定的類型，則傳回 **null**。||
| _options_|**物件**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _id_|**字串**|指定要用於識別新繫結物件的唯一名稱。如果針對 _id_ 參數未傳遞任何引數，則會自動產生 [Binding.id](../../reference/shared/binding.id.md)。||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**物件**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## 回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **addFromSelectionAsyn** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用途**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取代表由使用者指定之選取範圍的 [Binding](../../reference/shared/binding.md) 物件。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## 備註

將指定類型的繫結物件新增至 **Bindings** 集合，可由提供的 _id_ 識別。


 >**附註**  在 Excel 中，如果您呼叫在現有繫結之**Binding.id** 中傳遞的 **addFromSelectionAsync** 方法，則會使用該繫結的 [Binding.type](../../reference/shared/binding.type.md)，而且無法藉由指定 _bindingType_ 參數的不同值以變更其類型。如果您需要使用現有的 _id_ ，並變更 _bindingType_，請先呼叫 [Bindings.releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md) 方法，以釋出繫結，然後呼叫 **addFromSelectionAsync** 方法，以重新建立新類型的繫結。含


## 範例

使用 [Binding.id](../../reference/shared/binding.textbinding.md) ‘MyBinding’，將 **TextBinding** 新增至目前的選取範圍。


```js
function addBindingFromSelection() {
    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'MyBinding' }, 
        function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    );
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**可用於需求集合**|MatrixBindings, TableBindings, TextBindings|
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.1|在 Excel 增益集中，您可以針對包含表格式資料的儲存格範圍建立表格繫結 (傳遞 _bindingType_ 做為 **Office.BindingType.Table**)，即使該資料未新增至試算表做為表格 (藉由使用 **[插入]**  >  **[表格]**  >  **[表格]** 或 **[首頁]**  >  **[樣式]**  >  **[格式為表格]** 命令)。|
|1.1|新增支援 Access 內容增益集中的表格繫結。 |
|1.0|已導入|
