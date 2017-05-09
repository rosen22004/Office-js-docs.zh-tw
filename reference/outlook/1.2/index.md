# <a name="outlook-add-in-api-requirement-set-12"></a>Outlook 增益集 API 需求集合 1.2

適用於 Office 的 JavaScript API 的 Outlook 增益集 API 子集合包括物件、方法、屬性和事件，您可以用於 Outlook 增益集中。

> **附註**：本文件適用於[需求集合](../tutorial-api-requirement-sets.md)，而不是最新的需求集合。 

## <a name="whats-new-in-12"></a>1.2 中的新增功能？

需求集合 1.2 包括[需求集合 1.1](../1.1/index.md) 的所有功能。此會新增增益集的能力，可在郵件的主旨或內文以使用者的游標插入文字。

### <a name="change-log"></a>變更記錄

- 已新增 [Office.context.mailbox.item.setSelectedDataAsync](Office.context.mailbox.item.md#setSelectedDataAsync)：以非同步方式將資料插入至郵件的內文或主旨中。
- 已修改 [Office.context.mailbox.item.displayReplyAllForm](Office.context.mailbox.item.md#displayReplyAllForm)：已將 `attachments` 屬性新增至 `formData` 參數。
- 已修改 [Office.context.mailbox.item.displayReplyForm](Office.context.mailbox.item.md#displayReplyForm)：已將 `attachments` 屬性新增至 `formData` 參數。

## <a name="additional-resources"></a>其他資源

- [Outlook 增益集](../../../docs/outlook/outlook-add-ins.md)
- [Outlook 增益集程式碼範例](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [開始使用](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
