
# BindingDataChangedEventArgs 物件
提供引發 [DataChanged](../../reference/shared/binding.bindingdatachangedevent.md) 事件之繫結的相關資訊。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Word|
|**上次變更於 BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## 成員


**屬性**


|**名稱**|**說明**|
|:-----|:-----|
|[繫結](../../reference/shared/binding.bindingdatachangedeventargs.binding.md)|取得代表引發 [DataChanged](../../reference/shared/binding.md) 事件之繫結的 **Binding** 物件。|
|[類型](../../reference/shared/binding.bindingdatachangedeventargs.type.md)|取得可識別所引發事件類型的 [EventType](../../reference/shared/eventtype-enumeration.md) 列舉值。|

## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.1|新增對 Access 增益集中該事件的支援。|
|1.0|已導入|
