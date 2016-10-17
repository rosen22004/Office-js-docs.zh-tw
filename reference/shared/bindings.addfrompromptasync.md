
# <a name="bindings.addfrompromptasync-method"></a>Bindings.addFromPromptAsync 方法
 顯示 UI，讓使用者可指定要繫結的選取範圍。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|不在集合中|
|**上次變更**|1.1|

```
_bindingsObj.addFromPromptAsync(bindingType [, options], callback);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|指定要建立的繫結物件類型。必要。如果選取的物件無法強制轉型至指定的類型，則傳回 **null**。||
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _id_|**string**|指定要用於識別新繫結物件的唯一名稱。如果針對 _id_ 參數未傳遞任何引數，則會自動產生 [Binding.id](../../reference/shared/binding.id.md)。||
| _promptText_|**string**|指定要在提示 UI 中顯示的字串，告知使用者要選取的內容。限制為 200 個字元。如果未傳遞任何 _promptText_ 引數，則會顯示「請進行選擇」。||
| _sampleData_|[TableData](../../reference/shared/tabledata.md)|指定提示 UI 中顯示的範例資料表格，做為可由增益集繫結之欄位 (欄) 種類的範例。**TableData** 物件中提供的標頭可指定欄位選擇 UI 中所使用的標籤。選擇性。**附註：**此參數僅用於 Access 的增益集。如果在 Excel 增益集中呼叫方法時提供此參數，會予以忽略。||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **addFromPromptAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取代表由使用者指定之選取範圍的 [Binding](../../reference/shared/binding.md) 物件。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

將指定類型的繫結物件新增至 [Bindings](../../reference/shared/bindings.bindings.md) 集合，可由提供的 _id_ 識別。如果無法繫結指定的選取範圍，此方法將會失敗。


## <a name="example"></a>範例




```js
function addBindingFromPrompt() {

    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'MyBinding', promptText: 'Select text to bind to.' }, function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|不在集合中|
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 的支援。|
|1.1|在 Excel 增益集中，您可以針對包含表格式資料的儲存格範圍建立表格繫結 (傳遞 _bindingType_ 做為 **Office.BindingType.Table**)，即使該資料在 Excel UI 中未新增至試算表做為表格 (藉由使用 **[插入]**  >  **[表格]**  >  **[表格]** 或 **[首頁]**  >  **[樣式]**  >  **[格式為表格]** 命令)。|
|1.1|新增支援 Access 內容增益集中的表格繫結。 |
|1.1|新增支援在 Excel 增益集中繫結至矩陣資料做為表格繫結。|
|1.0|已導入|
