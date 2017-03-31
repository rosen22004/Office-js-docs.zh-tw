
# <a name="bindingbindingdatachanged-event"></a>Binding.bindingDataChanged 事件
繫結內的資料變更時，就會發生。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Word|
|**上次變更於 BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## <a name="remarks"></a>備註

若要新增繫結之 **BindingDataChanged** 事件的事件處理常式，請使用 [Binding](../../reference/shared/binding.addhandlerasync.md) 物件的 **addHandlerAsync** 方法。事件處理常式會收到 [BindingDataChangedEventArgs](../../reference/shared/binding.bindingdatachangedeventargs.md) 類型的引數。


## <a name="example"></a>範例




```js
function addEventHandlerToBinding() {
    Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
}

function onBindingDataChanged(eventArgs) {
    write("Data has changed in binding: " + eventArgs.binding.id);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|BindingEvents|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄

|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.1|新增對 Access 增益集中該事件的支援。|
|1.0|已導入|
