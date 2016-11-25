
# <a name="document.selectionchanged-event"></a>Document.SelectionChanged 事件
文件中的選取項目變更時，就會發生。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、PowerPoint、Word|
|**導入在**|1.1|

```
Office.EventType.DocumentSelectionChanged
```

## <a name="remarks"></a>備註

若要新增文件之 **SelectionChanged** 事件的事件處理常式，請使用 [Document](../../reference/shared/document.addhandlerasync.md) 物件的 **addHandlerAsync** 方法。


## <a name="example"></a>範例




```
function addEventHandlerToDocument() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}

function MyHandler(eventArgs) {
    doSomethingWithDocument(eventArgs.document);
}

```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.0|已導入|