
# <a name="documentselectionchangedeventargs-object"></a>DocumentSelectionChangedEventArgs 物件
提供引發 [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) 事件的文件相關資訊。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、PowerPoint、Word|
|**已新增於**|1.1|

```

```


## <a name="members"></a>成員


**屬性**


|**名稱**|**描述**|
|:-----|:-----|
|[document](../../reference/shared/document.selectionchangedeventargs.document.md)|取得代表引發 **SelectionChanged** 事件之文件的 **Document** 物件。|
|[type](../../reference/shared/document.selectionchangedeventargs.type.md)|取得可識別所引發事件類型的 **EventType** 列舉值。|

## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

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