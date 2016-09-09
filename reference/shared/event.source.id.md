

# event.source.id
取得觸發呼叫此函數的控制項識別碼。

****

|||
|:-----|:-----|
|**主應用程式：**Outlook|**增益集類型︰**Outlook|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|信箱|
|**上次變更於信箱**|1.3|
|**適用的 Outlook 模式**|讀取和撰寫|



```js
event.source.id;
```


## 傳回值

觸發呼叫此函數的控制項識別碼。識別碼來自資訊清單。


## 支援詳細資料


下表中的大寫 Y，表示在相對應的 Outlook 主應用程式中支援此屬性。空白儲存格表示 Outlook 主應用程式不支援此屬性。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。

 **重要事項：**增益集命令及與其相關聯的 API 目前只能在 Windows 桌面上，於 [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) 的 Outlook 中運作。


**支援的主應用程式 (依平台排序)**

| |**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**裝置適用的 OWA**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|||

|||
|:-----|:-----|
|**可用於需求集合**|信箱|
|**最低權限等級**|[ReadWriteItem](../../docs/outlook/understanding-outlook-add-in-permissions.md)|
|**增益集類型**|Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.3|已導入|
