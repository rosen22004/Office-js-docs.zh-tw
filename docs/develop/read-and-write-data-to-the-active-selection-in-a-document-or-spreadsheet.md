
# 在文件或試算表中的作用選取範圍內讀取和寫入資料

[Document](../../reference/shared/document.md) 物件公開的方法可讓您讀取和寫入至文件或試算表中使用者目前的選取範圍。若要這樣做，**Document** 物件可提供 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法。本主題也會說明如何讀取、寫入和建立事件處理常式，來偵測使用者選取範圍的變更。


  **getSelectedDataAsync** 方法只適用使用者目前的選取範圍。如果您需要保存文件中的選取範圍，使相同的選取範圍可來讀取和寫入執行增益集的各個工作階段，您必須使用 [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) 方法來加入繫結 (或使用 [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx) 物件的另一個 "addFrom" 方法來建立繫結)。如需建立繫結至文件的區域，然後讀取及寫入繫結的詳細資訊，請參閱[繫結至文件或試算表中的區域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。


### 讀取選取的資料


下列範例顯示如何藉由使用 [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) 方法，從文件中的選取範圍取得資料。


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在這個範例中，第一個 _coercionType_ 參數會指定為 **Office.CoercionType.Text** (您也可以使用常值字串 `"text"` 來指定這個參數)。這表示可以從回撥函式的 [asyncResult](../../reference/shared/asyncresult.status.md) 參數中取得的 [AsyncResult](../../reference/shared/asyncresult.md) 物件的 _value_ 屬性將傳回包含文件中選取文字的 **string**。指定不同的強制型轉類型將會導致不同的值。[Office.CoercionType](../../reference/shared/coerciontype-enumeration.md) 是可用的強制型轉類型值的列舉。**Office.CoercionType.Text** 評估為字串 "text"。


 >**提示：**   **何時應使用矩陣與資料表 coercionType 進行資料存取？** 如果您需要選取的表格式資料在加入列及欄時動態成長，而您必須處理表格標題，應該使用表格式資料類型 (藉由將 **getSelectedDataAsync** 方法的 _coercionType_ 參數指定為 `"table"` 或 **Office.CoercionType.Table**)。 於資料結構內加入列及欄在資料表和矩陣資料中支援，但附加列及欄僅對表格資料支援。 如果您不打算加入列及欄，並且您的資料並不需要標頭功能，那麼您應該使用矩陣資料類型 (藉由將 **getSelecteDataAsync** 方法的 _coercionType_ 參數指定為 `"matrix"` 或 **Office.CoercionType.Matrix**)，提供與資料互動的更簡單模型。

傳遞到函式作為第二個 _callback_ 參數的匿名函式會在 **getSelectedDataAsync** 作業完成時執行。會使用單一參數 _asyncResult_ 來呼叫函式，其中包含結果和呼叫的狀態。如果呼叫失敗，[AsyncResult](../../reference/shared/asyncresult.context.md) 物件的 **error** 屬性會提供 [Error](../../reference/shared/error.md) 物件的存取。您可以檢查 [Error.name](../../reference/shared/error.name.md) 和 [Error.message](../../reference/shared/error.message.md) 屬性的值，以判斷 set 作業失敗的原因。否則，會顯示文件中選取的文字。

[AsyncResult.status](../../reference/shared/asyncresult.error.md) 屬性用於 **if** 陳述式來測試呼叫是否成功。[Office.AsyncResultStatus](../../reference/shared/asyncresultstatus-enumeration.md) 是可用的 **AsyncResult.status** 屬性值的列舉。**Office.AsyncResultStatus.Failed** 評估為字串 "failed" (並且，同樣也可以指定為該常值字串)。


### 將資料寫入至選取範圍


下列範例會示範如何設定選取範圍以顯示 "Hello World!"。


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

為 _data_ 參數傳入不同的物件類型會有不同的結果。結果根據目前在文件中選取的項目、主控增益集的應用程式為何，以及傳入的資料是否可以強制轉型為目前的選取範圍而定。

傳遞至 [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) 方法做為 _callback_ 參數的匿名函式會在非同步呼叫完成時執行。當您使用 **setSelectedDataAsync** 方法將資料寫入至選取範圍，回撥的 _asyncResult_ 的參數只能夠提供呼叫狀態以及呼叫失敗時 [Error](../../reference/shared/error.md) 物件的存取。

 **附註：**隨著 Excel 2013 SP1 發行與 Excel Online 的對應組建開始，您現在可以[在寫入表格至目前選取範圍時設定格式設定](../../docs/excel/format-tables-in-add-ins-for-excel.md)。


### 偵測選取範圍中的變更


下列範例顯示如何藉由使用 [Document.addHandlerAsync](../../reference/shared/document.addhandlerasync.md) 方法來偵測選取範圍中的變更，以便為文件上的 [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) 事件加入事件處理常式。


```
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

第一個 _eventType_ 參數會指定要訂閱的事件名稱。對這個參數傳遞字串 `"documentSelectionChanged"` 相當於傳遞 **Office.EventType** 列舉的 [Office.EventType.DocumentSelectionChanged](../../reference/shared/eventtype-enumeration.md) 的事件類型。

傳遞至函式作為第二個 _handler_ 參數的 `myHander()` 函式，是當文件上的選取範圍變更時執行的事件處理常式。會使用單一參數 _eventArgs_ 來呼叫函式，其中將包含非同步作業完成時 [DocumentSelectionChangedEventArgs](../../reference/shared/document.selectionchangedeventargs.md) 物件的參考。您可以使用 [DocumentSelectionChangedEventArgs.document](../../reference/shared/document.selectionchangedeventargs.document.md) 屬性來存取引發事件的文件。


 >**附註：**您可以藉由再次呼叫 **addHandlerAsync** 方法，為指定的事件加入多個事件處理常式，並為 _handler_ 參數傳入其他的事件處理常式函式。只要每個事件處理常式函式的名稱是唯一的，這將正確運作。


### 停止偵測選取範圍中的變更


下列範例示範如何藉由呼叫 [document.removeHandlerAsync](../../reference/shared/document.selectionchanged.event.md) 方法來停止接聽 [Document.SelectionChanged](../../reference/shared/document.removehandlerasync.md) 事件。


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

傳遞作為第二個 _handler_ 參數的 `myHandler` 函式名稱，會指定將從 **SelectionChanged** 事件移除的事件處理常式。


 >**重要：**如果在呼叫 _removeHandlerAsync_ 方法時省略選用的 **handler** 參數，會移除所有指定的 _eventType_ 的事件處理常式。

