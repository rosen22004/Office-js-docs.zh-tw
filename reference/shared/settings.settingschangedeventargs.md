# <a name="settings.settingschangedeventargs-object"></a>Settings.settingschangedeventargs 物件
提供引發 [settingsChanged 事件](settings.settingschangedevent.md)的設定相關資訊。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel |
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|設定|
|**上次變更於**|1.0|

```js
Office.EventType.SettingsChanged
```

## <a name="members"></a>成員

**屬性**

|**名稱**|**描述**|
|:-----|:-----|
|**[settings](settings.settingschangedeventargs.setting.md)**|取得代表引發 settingsChanged 事件之設定的 **Settings** 物件。|
|**[type](settings.settingschangedeventargs.type.md)**|取得可識別所引發事件類型的 **EventType** 列舉值。|

## <a name="remarks"></a>備註

若要新增 **settingsChanged** 事件的事件處理常式，請使用 [Settings](settings.addhandlerasync.md) 物件的 **addHandlerAsync** 方法。

僅在您增益集的指令碼呼叫 **Settings.saveAsync** 方法，將設定的記憶體內部複本保存至文件檔案時，**settingsChanged** 事件才會啟動。呼叫 **Settings.set** 或 [Settings.remove](settings.set.md) 事件時，不會觸發 [settingsChanged](settings.remove.md) 事件。

當您的增益集用於共用 (共同編寫) 文件時，**settingsChanged** 事件可讓您在多位使用者同時嘗試儲存設定時，處理潛在衝突。


 >**重要**：您的增益集程式碼可在增益集透過任何 Excel 用戶端執行時，註冊 **settingsChanged** 事件的處理常式，但僅會在透過 Excel Online 開啟的試算表載入增益集，_且_有多位使用者編輯試算表 (共同編寫) 時，啟動此事件。因此，實際上僅在 Excel Online 中的共同編寫情節下，支援 **settingsChanged** 事件。



## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||


|||
|:-----|:-----|
|**可用於需求集合**|Settings|
|**最低權限等級**|Restricted|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄

|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|
