
# <a name="documentactiveviewchanged-event"></a>Document.ActiveViewChanged 事件
使用者變更文件目前的檢視時，就會發生。

|||
|:-----|:-----|
|**主應用程式︰**|PowerPoint|
|**導入在**|1.1|

```
Office.EventType.ActiveViewChanged
```


## <a name="remarks"></a>備註

若要新增文件之 **ActiveViewChanged** 事件的事件處理常式，請使用 [Document](../../reference/shared/document.addhandlerasync.md) 物件的 **addHandlerAsync** 方法。事件處理常式會收到 [ActiveViewChangedEventArgs](../../reference/shared/document.activeviewchangedeventargs.md) 類型的引數。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for Mac**|**Office for iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y||Y|Y|

|||
|:-----|:-----|
|**導入在**|1.1|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|
