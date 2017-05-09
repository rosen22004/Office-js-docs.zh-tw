# <a name="outlook-add-in-api-requirement-set-13"></a>Outlook 增益集 API 需求集合 1.3

適用於 Office 的 JavaScript API 的 Outlook 增益集 API 子集合包括物件、方法、屬性和事件，您可以用於 Outlook 增益集中。

> **附註**：本文件適用於[需求集合](../tutorial-api-requirement-sets.md)，而不是最新的需求集合。 

## <a name="whats-new-in-13"></a>1.3 中的新增功能？

需求集合 1.3 包括[需求集合 1.2](../1.2/index.md) 的所有功能。其中新增下列功能。

- 為[增益集命令](../../../docs/outlook/add-in-commands-for-outlook.md)新增支援。
- 新增能夠將撰寫中項目儲存或關閉的能力。
- 已增強 [Body](Body.md) 物件，允許增益集取得或設定整個內文。
- 已新增可在 EWS 和 REST 格式間轉換 ID 的轉換方法。
- 已新增可將通知訊息新增至項目上 [資訊] 列的能力。

### <a name="change-log"></a>變更記錄

- 已新增 [Body.getAsync](Body.md#getAsync)：以指定的格式傳回目前的本文。
- 已新增 [Body.setAsync](Body.md#setAsync)：將內文全文取代為指定文字。
- 已新增 [Office.context.officeTheme](Office.context.md#officeTheme)：提供 Office 佈景主題色彩的存取。
- 已新增 [事件](Event.md) 物件︰做為參數傳遞至 Outlook 增益集中的無 UI 命令函數。用來發出處理完成的訊號。
- 已新增 [Office.context.mailbox.item.close](Office.context.mailbox.item.md#close)：關閉目前正在撰寫的項目。
- 已新增 [Office.context.mailbox.item.saveAsync](Office.context.mailbox.item.md#saveAsync)：以非同步方式儲存項目。
- 已新增 [Office.context.mailbox.item.notificationMessages](Office.context.mailbox.item.md#notificationMessages)：取得項目的通知訊息。
- 已新增 [Office.context.mailbox.convertToEwsId](Office.context.mailbox.md#convertToEwsId)：將 REST 的項目 ID 轉換為 EWS 格式。
- 已新增 [Office.context.mailbox.convertToRestId](Office.context.mailbox.md#convertToRestId)：將 EWS 的項目 ID 轉換為 REST 格式。
- 已新增 [Office.MailboxEnums.ItemNotificationMessageType](Office.MailboxEnums.md#ItemNotificationMessageType)：指定約會或郵件的通知訊息類型。
- 已新增 [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#RestVersion)：指定與 REST 格式的項目 ID 對應的 REST API 版本。
- 已新增 [NotificationMessages](NotificationMessages.md) 物件：提供方法來存取 Outlook 增益集的通知訊息。
- 已新增 [NotificationMessageDetails](simple-types.md#NotificationMessageDetails) 類型：由 `NotificationMessages.getAllAsync` 方法所傳回。

## <a name="additional-resources"></a>其他資源

- [Outlook 增益集](../../../docs/outlook/outlook-add-ins.md)
- [Outlook 增益集程式碼範例](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [開始使用](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
