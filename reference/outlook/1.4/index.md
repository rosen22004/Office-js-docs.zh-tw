# <a name="outlook-add-in-api-requirement-set-14"></a>Outlook 增益集 API 需求集合 1.4

適用於 Office 的 JavaScript API 的 Outlook 增益集 API 子集合包括物件、方法、屬性和事件，您可以用於 Outlook 增益集中。

> **附註**：本文件適用於[需求集合](../tutorial-api-requirement-sets.md)，而不是最新的需求集合。

## <a name="whats-new-in-14"></a>1.4 中的新增功能？

需求集合 1.4 包括[需求集合 1.3](../1.3/index.md) 的所有功能。其已將存取權新增至 `Office.ui` 命名空間。

### <a name="change-log"></a>變更記錄

- 已新增 [Office.context.ui.displayDialogAsync](../../shared/officeui.displaydialogasync.md)：在 Office 主應用程式中顯示對話方塊。
- 已新增 [Office.context.ui.messageParent](../../shared/officeui.messageparent.md)：從對話方塊將訊息傳遞至其父系/opener 頁面。
- 已新增 [Dialog](../../shared/officeui.dialog.md) 物件：呼叫 `displayDialogAsync` 方法時傳回的物件。

## <a name="additional-resources"></a>其他資源

- [Outlook 增益集](../../../docs/outlook/outlook-add-ins.md)
- [Outlook 增益集程式碼範例](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [開始使用](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
