
# TextBinding 物件
表示文件中的繫結文字選取。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、PowerPoint、Project、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TextBindings|
|**已新增於**|1.0|

```
TextBinding
```


## 備註

**TextBinding** 物件從 [Binding](../../reference/shared/binding.id.md) 物件繼承 [id](../../reference/shared/binding.type.md) 屬性、[type](../../reference/shared/binding.getdataasync.md) 屬性、[getDataAsync](../../reference/shared/binding.setdataasync.md) 方法，以及 [setDataAsync](../../reference/shared/binding.md) 方法。其未實作任何其他屬性或本身的方法。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**可用於需求集合**|TextBindings|
|**最低權限等級**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.0|已導入|
