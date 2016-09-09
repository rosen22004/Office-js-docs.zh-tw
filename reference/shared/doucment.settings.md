
# 設定物件
代表以名稱/值對儲存在主文件中，工作窗格或內容增益集的自訂設定。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|設定|
|**上次變更於**|1.1|

```
Office.context.document.settings
```


## 成員


**方法**

|||
|:-----|:-----|
|名稱|描述|
|[addHandlerAsync](../../reference/shared/settings.addhandlerasync.md)|新增 **settingsChanged** 事件的事件處理常式。|
|[get](../../reference/shared/settings.get.md)|擷取指定的設定。|
|[refreshAsync](../../reference/shared/settings.refreshasync.md)|讀取保存在文件內的所有設定，並重新整理保留在記憶體中，這些設定的增益集複本。|
|[remove](../../reference/shared/settings.remove.md)|移除指定的設定。|
|[removeHandlerAsync](../../reference/shared/settings.removehandlerasync.md)|移除 **settingsChanged** 事件的事件處理常式。|
|[saveAsync](../../reference/shared/settings.saveasync.md)|儲存設定。|
|[set](../../reference/shared/settings.set.md)|設定或建立指定的設定。|

**事件**


|**名稱**|**描述**|
|:-----|:-----|
|[settingsChanged](../../reference/shared/settings.settingschangedevent.md)|設定變更時，就會發生。|

## 備註

使用 **Settings** 物件的方法所建立的設定，會按照每個增益集和每個文件而儲存。也就是說，它們只能用於建立它們的增益集，而且只能從儲存它們的文件中取用。

設定的名稱為**字串**，而值可以是**字串**、**數字**、**布林值**、**null**、**物件**或**陣列**。

**Settings** 物件會自動載入為 [Document](../../reference/shared/document.md) 物件的一部分，並且在增益集啟動時，藉由呼叫該物件的 [settings](../../reference/shared/document.settings.md) 屬性使用。新增或刪除設定以儲存文件中的設定之後，開發人員負責呼叫 [saveAsync](../../reference/shared/settings.saveasync.md) 方法。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**iPad 版 Office**|
|:-----|:-----|:-----|:-----|
|**Access**|Y|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|設定|
|**增益集類型**|內容、工作窗格|
|**程式庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>對於 <a href="7c4780cf-a779-4ac9-a362-c0bacae64a96.htm">addHandlerAsync</a> 和 <a href="735a255b-2a86-4b43-b1fa-e2a305815615.htm">removeHandlerAsync</a> 方法，已新增在 Access 內容增益集中，新增與移除 <span class="keyword">SettingsChanged</span> 事件的事件處理常式。</p></li><li><p>對於 <a href="aeac06dd-994e-4235-b208-1bd117395296.htm">get</a>、<a href="53a52c47-24b4-4d2d-b840-fe1b242cd795.htm">refreshAsync</a>、<a href="a92446bf-de65-45bd-8412-36ea8e77c5a2.htm">remove</a>、<a href="7147c221-937c-477c-98a6-f59d6200c27b.htm">saveAsync</a> 和 <a href="4e2c9758-953e-41e8-aca6-d8daf764a584.htm">set</a>，新增支援 Access 內容增益集中的自訂設定。</p></li></ul>|
|1.0|已導入|

