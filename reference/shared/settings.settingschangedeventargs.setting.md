

# <a name="settingschangedeventargs.settings-property"></a>SettingsChangedEventArgs.settings 屬性
取得代表引發 **settingsChanged** 事件之設定的 **Settings** 物件。

|||
|:-----|:-----|
|**主應用程式︰**|Excel|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|設定|
|**上次變更於**|1.0|

```js
var mySettings = eventArgsObj.settings;
```


## <a name="return-value"></a>傳回值

代表引發 [settingsChanged](../../reference/shared/document.settings.md) 事件之設定的 [Settings](../../reference/shared/settings.settingschangedevent.md) 物件。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||

|||
|:-----|:-----|
|**可用於需求集合**|Settings|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|