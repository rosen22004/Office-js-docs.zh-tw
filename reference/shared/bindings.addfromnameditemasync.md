
# Bindings.addFromNamedItemAsync 方法
將繫結加入至文件中的具名項目。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**上次變更**|1.1|

```
Office.context.document.bindings.addFromNamedItemAsync(itemName, bindingType [, options], callback);
```


## 參數



|**名稱**|**類型	**|**說明**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _itemName_|**string**|具名項目的名稱。必要。||
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|指定要建立的繫結物件類型。必要。如果選取的物件無法強制轉型至指定的類型，則傳回 **null**。||
| _options_|**物件**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _id_|**字串**|指定要用於識別新繫結物件的唯一名稱。如果針對 _id_ 參數未傳遞任何引數，則會自動產生 [Binding.id](../../reference/shared/binding.id.md)。||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**物件**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## 回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **addFromNamedItemAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用途**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取代表指定之具名項目的 [Binding](../../reference/shared/binding.md) 物件。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## 備註

 **對於 Excel**，_itemName_ 參數可以指具名的範圍或表格。

依預設，在 Excel 中新增表格會針對您新增的第一個表格指派名稱「Table1」，針對您新增的第二個表格指派名稱 「Table2」，依此類推。若要在 Excel UI 中為表格指派有意義的名稱，請使用功能區的 **[資料表工具] | [設計]** 索引標籤上的**表格名稱**屬性。


 >**附註：**在 Excel 中，將表格指定為具名項目時，您必須指定完整名稱，才能使用此格式將工作表名稱加入表格名稱：`"Sheet1!Table1"`

 **對於 Word**，_itemName_ 參數是指 **RTF 文字**內容控制項的**標題**屬性。(您無法繫結至 **RTF 文字**內容控制項以外的內容控制項。)

依預設，內容控制項沒有指派的**標題**值。若要在 Word UI 中指派有意義的名稱，從功能區 **[開發人員]** 索引標籤上的**控制項**群組插入**RTF 文字**內容控制項之後，請使用**控制項**群組中的 **Properties** 命令，以顯示 **[內容控制項屬性]** 對話方塊。然後將內容控制項的**標題**屬性設為您要從程式碼參照的名稱。


 >**附註**  在 Word 中，如果有多個具有相同**標題**屬性值 (名稱) 的**RTF 文字**內容控制項，而且您嘗試使用此方法 (藉由將其名稱指定為 _itemName_ 參數) 繫結至其中一個內容控制項，則作業將會失敗。


## 範例

下列範例會將繫結新增至 Excel 中的 `myRange` 具名項目做為 “matrix” 繫結，並指派繫結的 [id](../../reference/shared/binding.id.md) 作為 `myMatrix`。


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

下列範例會將繫結新增至 Excel 中的 `Table1` 具名項目做為 “table” 繫結，並指派繫結的 **id** 作為 `myTable`。




```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("Table1", "table", {id:'myTable'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

下列範例會在 Word 中建立文字繫結至名稱為 `"FirstName"`的 RTF 文字內容控制項、指派 **id**`"firstName"`，然後顯示該資訊。




```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
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

## 請參閱



#### 其他資源


[繫結至文件或試算表中的區域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md#add-a-binding-to-a-named-item)
