# <a name="outlook-add-in-api-requirement-set-15"></a>Outlook 增益集 API 需求集合 1.5

適用於 Office 的 JavaScript API 的 Outlook 增益集 API 子集合包括物件、方法、屬性和事件，您可以用於 Outlook 增益集中。

## <a name="whats-new-in-15"></a>1.5 中的新增功能？

需求集合 1.5 包括[需求集合 1.4](../1.4/index.md) 的所有功能。其中新增下列功能。

- 新增對[可釘選的工作窗格](../../../docs/outlook/manifests/pinnable-taskpane.md)的支援。
- 新增對呼叫 [REST API](../../../docs/outlook/use-rest-api.md) 的支援。
- 新增將附件標示為內嵌的功能。
- 新增將工作窗格或對話方塊關閉的的功能。

### <a name="change-log"></a>變更記錄

- 已新增 [Office.context.mailbox.addHandlerAsync](Office.context.mailbox.md#addHandlerAsync)：新增支援事件的事件處理常式。
- 已新增 [Office.EventType](Office.md#EventType)：指定與事件處理常式相關聯的事件。
- 已新增 [Office.context.mailbox.restUrl](Office.context.mailbox.md#restUrl)：取得此電子郵件帳戶的 REST 端點的 URL。
- 已修改 [Office.context.mailbox.getCallbackTokenAsync](Office.context.mailbox.md#getCallbackTokenAsync)：已新增這個含簽章的新版本方法 (`getCallbackTokenAsync([options], callback)`)。原始的版本維持不變且仍可使用。
- 已新增 [Office.context.ui.closeContainer](Office.context.ui.md#closeContainer)： 
- 已修改 [Office.context.mailbox.item.addFileAttachmentAsync](Office.context.mailbox.item.md#addFileAttachmentAsync)：在名為 `isInline` 的 `options` 字典中的新值，是用來指定影像在訊息內文中以內嵌使用。
- 已修改 [Office.context.mailbox.item.displayReplyAllForm](Office.context.mailbox.item.md#displayReplyAllForm)：在名為 `isInline` 的 `formData.attachments` 字典中的新值，是用來指定影像在訊息內文中以內嵌使用。
- 已修改 [Office.context.mailbox.item.displayReplyForm](Office.context.mailbox.item.md#displayReplyForm)：在名為 `isInline` 的 `formData.attachments` 字典中的新值，是用來指定影像在訊息內文中以內嵌使用。

## <a name="additional-resources"></a>其他資源

- [Outlook 增益集](../../../docs/outlook/outlook-add-ins.md)
- [Outlook 增益集程式碼範例](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [開始使用](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
