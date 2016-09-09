
# 比較 Outlook 增益集在 Mac 版 Outlook 與其他 Outlook 主機中的支援

您可以在其他主機 (包括 Windows 版 Outlook、裝置的 OWA 及 Outlook Web App) 中使用與 Mac 版 Outlook 相同的方式建立並執行 Outlook 增益集，而不需要為每個主機自訂 JavaScript。從增益集對適用於 Office 的 JavaScript API 的相同呼叫通常運作方式相同，除了下表所述的區域以外。

 >**附註**  Mac 版 Outlook 僅在 Outlook 讀取模式中支援適用於 Office 的 JavaScript API。

|**適用範圍**|**Windows 版 Outlook、裝置的 OWA、Outlook Web App**|**Mac 版 Outlook**|
|:-----|:-----|:-----|
|office.js 的支援版本和 Office 增益集資訊清單結構|Office.js 和結構描述 1.1 版中的所有 API。|<ul><li>僅限可在讀取模式中使用的 API。可以啟動使用 office.js 1.1 版中新的和擴充 API 的增益集，但撰寫模式的 API 無法在 Mac 版 Outlook 上正確執行。 </li><li>結構描述 1.1 版。</li></ul>|
|週期性約會系列的執行個體|<ul><li>可以取得項目 ID 以及主約會或週期性系列的約會執行個體的其他屬性。 </li><li>可以使用 [mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md#displayappointmentformitemid) 以顯示執行個體或週期性系列的主圖形。</li></ul>|<ul><li>可以取得項目 ID 及主約會的其他屬性，但無法取得週期性系列執行個體的屬性。</li><li>可以顯示週期性系列的主約會。若沒有項目 ID，無法顯示週期性系列的執行個體。</li></ul>|
|約會出席者的收件者類型|可以使用 [EmailAddressDetails.recipientType](../../reference/outlook/simple-types.md) 來識別某位出席者的收件者類型。|**EmailAddressDetails.recipientType** 會針對約會出席者傳回**未定義**。|
|主機的版本字串 |[diagnostics.hostVersion](../../reference/outlook/Office.context.mailbox.diagnostics.md) 所傳回的版本字串格式會依主機的實際類型而定。例如︰<ul><li>Windows 版 Outlook15.0.4454.1002</li><li>Outlook Web App︰15.0.918.2</li></ul>|在 Mac 版 Outlook 上由 **Diagnostics.hostVersion** 所傳回的版本字串範例：15.0 (140325)|
|項目的自訂屬性|如果網路當機，增益集仍可存取快取的自訂屬性。|因為 Mac 版 Outlook 不會快取自定屬性，如果網路連線中斷，增益集就不能存取它們。|
|附件詳細資料|[AttachmentDetails](../../reference/outlook/Office.context.mailbox.md) 物件中的內容類型及附件名稱依主機類型而定：<ul><li>JSON 範例 <b>AttachmentDetails.contentType</b>: <b>"contentType": "image/x-png"</b>。 </li><li><b>AttachmentDetails.name</b> 不包含任何副檔名。例如，如果附件是具有以下主旨的郵件：「RE:Summer activity」，則代表附件名稱的 JSON 物件為 <b>「name」：「RE:Summer activity」</b>。</li></ul>|<ul><li>JSON 範例 <b>AttachmentDetails.contentType</b>: <b>"contentType": "image/png"</b></li><li><b>AttachmentDetails.name</b> 一律包含副檔名。郵件項目的附件具有 .eml 副檔名，而約會具有 .ics 副檔名。例如，如果附件為具有以下主旨的電子郵件：「RE:Summer activity」，則代表附件名稱的 JSON 物件為 <b>「name」：「RE:Summer activity.eml」</b>。</li></ul>|
|在 **dateTimeCreated** 及 **dateTimeModified** 屬性中代表時區的字串|範例：2014 年 3 月 13 日星期四 14:09:11 GMT+0800 (中國標準時間)|範例：2014 年 3 月 13 日星期四 14:09:11 GMT+0800 (CST)|
|**dateTimeCreated** 及 **dateTimeModified** 的時間精確度|如果增益集使用下列程式碼，則精確度會高達毫秒。<br/><pre lang="javascript">JSON.stringify(Office.context.mailbox.item, null, 4)；</pre>|精確度最高僅達秒。|

## 其他資源



- [部署和安裝 Outlook 增益集以進行測試](../outlook/testing-and-tips.md)
    
