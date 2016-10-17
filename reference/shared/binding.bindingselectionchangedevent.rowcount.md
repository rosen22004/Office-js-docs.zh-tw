
# <a name="bindingselectionchangedeventargs.rowcount-property"></a>BindingSelectionChangedEventArgs.rowCount 屬性
取得選取的資料列數目。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Word|
|**上次變更於**|1.1|

```
var rwCount = eventArgsObj.rowCount;
```


## <a name="return-value"></a>傳回值

選取的資料列數目。如果選取單一儲存格，則會傳回 1。


## <a name="remarks"></a>備註

如果使用者選取非連續的選取範圍，則會傳回繫結內最後一個連續選取的計數。 

對於 Word 而言，此屬性僅適用於 [BindingType](../../reference/shared/bindingtype-enumeration.md) "table" 的繫結。如果繫結是 "matrix" 類型，則會傳回 **null**。此外，如果表格包含合併的儲存格，則呼叫會失敗，因為表格結構必須針對此屬性統一，才可正常運作。


## <a name="example"></a>範例

下列範例會針對 [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) 事件，將事件處理常式新增至帶有 `myTable`[id](../../reference/shared/binding.id.md) 的繫結。當使用者變更選取範圍時，處理常式會顯示選取範圍中第一個儲存格的座標，以及所選取的資料列和資料欄數目。


```js
function addSelectionHandler() {
    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addHandlerAsync("bindingSelectionChanged", myHandler);
    });
}

// Display selection start coordinates and row/column count.
function myHandler(bArgs) {
    write("Selection start row/col: " + bArgs.startRow + "," + bArgs.startColumn);
    write("Selection row count: " + bArgs.rowCount);
    write("Selection col count: " + bArgs.columnCount);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援|
|1.1|您現在可以針對 Access 之內容增益集的  **SelectionChanged** 事件，新增並移除事件處理常式。|
|1.0|已導入|
