# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook 增益集 API 需求集合 1.1

適用於 Office 的 JavaScript API 的 Outlook 增益集 API 子集合包括物件、方法、屬性和事件，您可以用於 Outlook 增益集中。

> **附註**：本文件適用於[需求集合](../tutorial-api-requirement-sets.md)，而不是最新的需求集合。 

## <a name="whats-new-in-11"></a>1.1 中的新增功能

需求集合 1.1 包括需求集合 1.0 的所有功能。此會新增能力，讓增益集可存取訊息和約會的內文，以及新增對目前項目進行修改的能力。

### <a name="change-log"></a>變更記錄

- 新增 [主體](Body.md) 物件︰提供在 Outlook 增益集中新增和更新項目內容的方法。
- 新增 [位置](Location.md) 物件︰提供在 Outlook 增益集中取得及設定會議地點的方法。
- 新增 [收件者](Recipients.md) 物件：提供在 Outlook 增益集中取得及設定約會或訊息收件者的方法。
- 新增 [主旨](Subject.md) 物件：提供在 Outlook 增益集中取得及設定約會或訊息主旨的方法。
- 新增 [時間](Time.md) 物件：提供在 Outlook 增益集中取得及設定會議開始或結束時間的方法。
- 已新增 [Office.context.mailbox.item.addFileAttachmentAsync](Office.context.mailbox.item.md#addFileAttachmentAsync)：將檔案新增至郵件或約會做為附件。
- 已新增 [Office.context.mailbox.item.addItemAttachmentAsync](Office.context.mailbox.item.md#addItemAttachmentAsync)：將 Exchange 項目，例如訊息，新增至郵件或約會做為附件。
- 已新增 [Office.context.mailbox.item.removeAttachmentAsync](Office.context.mailbox.item.md#removeAttachmentAsync)：移除來自郵件或約會的附件。
- 已新增 [Office.context.mailbox.item.body](Office.context.mailbox.item.md#body)：取得提供方法來管理項目本文的物件。
- 已新增 [Office.context.mailbox.item.bcc](Office.context.mailbox.item.md#bcc)：取得或設定郵件 [密件副本] 行上的收件者。
- 已新增 [Office.MailboxEnums.RecipientType](Office.MailboxEnums.md#RecipientType)：指定約會的收件者類型。

## <a name="additional-resources"></a>其他資源

- [Outlook 增益集](../../../docs/outlook/outlook-add-ins.md)
- [Outlook 增益集程式碼範例](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [開始使用](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
