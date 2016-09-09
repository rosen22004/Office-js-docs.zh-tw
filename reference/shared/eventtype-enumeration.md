
# EventType 列舉
指定引發的事件種類。由 an **EventName**_EventArgs_ 物件的  **type** 屬性所傳回。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、PowerPoint、Project、Word|
|**上次變更於 Selection**|1.1|

```js
Office.EventType
```


## 成員


**值**


|列舉|值|描述|
|:-----|:-----|:-----|
|Office.EventType.ActiveViewChanged|"documentActiveViewChanged"|已引發 [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) 事件。|
|Office.EventType.DocumentSelectionChanged|"documentSelectionChanged"|已引發 [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md) 事件。|
|Office.EventType.BindingSelectionChanged|"bindingSelectionChanged"|已引發 [Binding.BindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) 事件。|
|Office.EventType.BindingDataChanged|"bindingDataChanged"|已引發 [Binding.BindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md) 事件。|
|Office.EventType.DataNodeDeleted|"nodeDeleted"|已引發 [CustomXmlPart.dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md) 事件。|
|Office.EventType.DataNodeInserted|"nodeInserted"|已引發 [CustomXmlPart.dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md) 事件。|
|Office.EventType.DataNodeReplaced|"nodeReplaced"|已引發 [CustomXmlPart.dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md) 事件。|
|Office.EventType.SettingsChanged|"settingsChanged"|已引發 [Settings.settingsChanged](../../reference/shared/settings.settingschangedevent.md) 事件。|

## 備註


 >**附註**：Project 的增益集支援  **Office.EventType.ResourceSelectionChanged**、 **Office.EventType.TaskSelectionChanged** 和  **Office.EventType.ViewSelectionChanged** 事件類型。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此列舉。空白儲存格表示 Office 主應用程式不支援此列舉。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y||
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1| 針對新的 **Document.ActiveViewChanged** 事件，新增 Office.EventType.ActiveViewChanged 列舉。|
|1.0|已導入|
