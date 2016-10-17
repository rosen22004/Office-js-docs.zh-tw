

# <a name="office.select-method"></a>Office.select 方法
建立承諾，以依據傳入的選取器字串傳回繫結。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**上次變更於**|1.1|

```js
Office.select(str, onError);
```


## <a name="parameters"></a>參數


_str_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**字串**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;所要剖析的選取器字串，並為其建立承諾。

_onError_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**函數**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。選用。
    

## <a name="callback-value"></a>回呼值

傳遞至 _onError_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。如果作業失敗，請使用 [AsyncResult.error](../../reference/shared/asyncresult.error.md) 屬性存取 [Error](../../reference/shared/error.md) 物件，以提供錯誤的相關資訊。


## <a name="remarks"></a>備註

**Office.select** 方法可供存取 [Binding](../../reference/shared/binding.md) 物件承諾，以嘗試在其任何非同步方法被叫用時，傳回指定的繫結。

支援的格式："bindings# _bindingId_"，它會針對具有 `bindingId` 的[id](../../reference/shared/binding.id.md) 的繫結，傳回 **Binding** 物件。如需詳細資訊，請參閱[在 Office 增益集中進行非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md#asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings)，以及[繫結至文件或試算表中的區域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。


 >**附註**：如果 **select** 方法承諾成功傳回 **Binding** 物件，則該物件只會公開 [Binding](../../reference/shared/binding.md) 物件的下列四個方法︰[getDataAsync](../../reference/shared/binding.getdataasync.md)、[setDataAsync](../../reference/shared/binding.setdataasync.md)、[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) 和 [removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)。如果該承諾無法傳回 **Binding** 物件，可以使用 _onError_ 回呼來存取 [asyncResult.error](../../reference/shared/asyncresult.error.md) 物件，以取得詳細資訊。如果您需要呼叫的 **Binding** 物件成員，不是 **select** 方法傳回之 **Binding** 物件承諾所公開的四個方法，則改為使用 [getByIdAsync](../../reference/shared/bindings.getbyidasync.md) 方法，您可以使用 [Document.bindings](../../reference/shared/document.bindings.md) 屬性和 [Bindings.getByIdAsync](../../reference/shared/bindings.getbyidasync.md) 方法，以擷取 **Binding** 物件。


## <a name="example"></a>範例

下列程式碼範例會使用 **select** 方法，從 **Bindings** 集合擷取具有 **id** " `cities`" 的繫結，然後呼叫 [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) 方法，為繫結的 [dataChanged](../../reference/shared/binding.bindingdatachangedevent.md) 事件加入事件處理常式。


```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}
```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**可用於需求集合**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**最低權限等級**|[ReadDocument (ReadAllDocument for Open Office XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援|
|1.1|新增使用 **select** 方法傳回在 Access 內容增益集中建立的資料表繫結。|
|1.0|已導入|
