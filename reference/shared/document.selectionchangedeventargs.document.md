
# <a name="documentselectionchangedeventargs.document-property"></a>DocumentSelectionChangedEventArgs.document 屬性
取得代表引發 **SelectionChanged** 事件之文件的 **Document** 物件。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、Word|
|**已新增於**|1.1|




```js
var myDoc = eventArgsObj.document;
```


## <a name="return-value"></a>傳回值

代表引發 [SelectionChanged](../../reference/shared/document.md) 事件之文件的 [Document](../../reference/shared/document.selectionchanged.event.md) 物件。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

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
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.0|已導入|