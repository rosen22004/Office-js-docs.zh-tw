
# BindingSelectionChangedEventArgs.binding 屬性
取得代表引發 **SelectionChanged** 事件之繫結的 **Binding** 物件。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Word|
|**上次變更於**|1.1|

```
var myBinding = eventArgsObj.binding;
```


## 傳回值

代表引發 [SelectionChanged](../../reference/shared/binding.md) 事件之繫結的 [Binding](../../reference/shared/binding.bindingselectionchangedevent.md) 物件。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄





****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.1|您現在可以針對 Access 之內容增益集的  **SelectionChanged** 事件，新增並移除事件處理常式。|
|1.0|已導入|
