
# <a name="tablebinding.addcolumnsasync-method"></a>TableBinding.addColumnsAsync 方法
將資料欄和值加入表格中。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**上次變更於**|1.0|

```
bindingObj.addColumnsAsync(data [, options], callback);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _data_|**array** 或 [TableData](../../reference/shared/tabledata.md)|陣列的陣列 (“matrix”) 或包含資料之一或多列以加入表格的 **TableData** 物件。必要。||
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **addColumnsAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|因為沒有可擷取的物件或資料，所以一律傳回 **undefined**。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

若要新增指定資料或標頭的值的一或多個資料欄，請傳遞  **TableData** 物件做為 _data_ 參數。若要新增僅指定資料的一或多個資料欄，請傳遞陣列的陣列 (“matrix”) 做為 _data_ 參數。

**addColumnAsync** 作業的成功或失敗是不可部分完成。也就是整個新增欄作業必須成功，否則將會完全復原 (而且傳回回呼的 **AsyncResult.status** 屬性將會回報失敗)：


- 您傳遞做為_資料_引數之陣列中的每一列，必須與要更新的表格具有相同的列數。如果沒有，整個作業將會失敗。
    
- 陣列中的每一列和儲存格必須將該列或儲存格成功新增至新增欄中的表格。如果基於任何原因造成無法設定任何列或儲存格，則整個作業將會失敗。
    
- 如果您傳遞  **TableData** 物件做為資料引數，則標題列的數目必須符合要更新之表格的數目。
    
**Excel Online 的其他備註**

要傳遞給 **data** 參數之 _TableData_ 物件中的儲存格總數，在針對此方法的單一呼叫中，不可超過 20,000。


## <a name="example"></a>範例

下列範例將藉由傳遞 [TableData](../../reference/shared/binding.id.md) 物件做為 **addColumnsAsync** 方法的_資料_引數，將帶有三個資料列的單一資料欄新增至帶有 **id**`"myTable"` 的繫結表格。若要成功，要更新的表格必須有三個資料列。


```js
// Add a column to a binding of type table by passing a TableData object.
function addColumns() {
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```

下列範例將藉由傳遞陣列的陣列 (“matrix”) 做為 [addColumnsAsync](../../reference/shared/binding.id.md) 方法的_資料_引數，將帶有三個資料列的單一資料欄新增至帶有 **id**`myTable` 的繫結表格。若要成功，要更新的表格必須有三個資料列。




```js
// Add a column to a binding of type table by passing an array of arrays.
function addColumns() {
    var myTable = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|TableBindings|
|**最低權限等級**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.0|已導入|
