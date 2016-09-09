
# DocumentActiveViewChangedEventArgs.type 屬性
取得可識別所引發事件類型的 **EventType** 列舉值。

|||
|:-----|:-----|
|**主機︰**|PowerPoint|
|**已新增於**|1.1|

```
var myEventType = eventArgsObj.type;
```


## 傳回值

引發之事件的 [EventType](../../reference/shared/eventtype-enumeration.md)。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

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
|1.1|新增 iPad 版 Office 中對 PowerPoint 的支援。|
|1.1|已導入。|
