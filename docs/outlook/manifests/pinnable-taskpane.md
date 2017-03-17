# <a name="implement-a-pinnable-taskpane-in-outlook"></a>在 Outlook 中實作可釘選的工作窗格

增益集命令的[工作窗格](../add-in-commands-for-outlook.md#launching-a-task-pane) UX 圖案，會在開啟的訊息或約會項目右方，開啟垂直工作窗格，可讓增益集提供具有更多詳細互動的 UI (像是填入多個欄位等等)。在檢視訊息清單時，此工作窗格可在 [讀取窗格] 中顯示，可讓您快速處理特定訊息。

然而，在預設情況下，若使用者於 [讀取窗格] 中已開啟了特定訊息的增益集工作窗格，接著選取新訊息，則工作窗格便會自動關閉。針對需頻繁使用的增益集，使用者可能會偏好讓該窗格保持開啟，而無需針對各個訊息重新啟動增益集。使用可釘選工作窗格，您的增益集便能提供使用者該選項。

> **注意**：可釘選的工作窗格目前僅 Outlook 2016 for Windows 提供支援 (組建 7668.2000 或更新版本適用於目前或 Office 測試人員通道中的使用者，組建 7900.xxxx 或更新版本適用於順延通道中的使用者)。

## <a name="support-taskpane-pinning"></a>支援工作窗格釘選

第一個步驟是新增釘選支援，可在增益集[資訊清單](./manifests.md)中完成。方式是將 [SupportsPinning](../../../reference/manifest/action.md#supportspinning) 元素新增至 `Action` 元素，其說明了工作窗格按鈕。

`SupportsPinning` 元素是在 VersionOverrides 1.1 版結構描述中定義，因此您需要加入可同時適用於 1.0 及 1.1 版的 [VersionOverrides](../../../reference/manifest/versionoverrides.md) 元素。

> **注意：**如果您打算將 Outlook 增益集[發佈](../../publish/publish.md)至 Office 市集，為了能通過 [Office 市集驗證 ](https://msdn.microsoft.com/en-us/library/jj220035.aspx)，在您使用 **SupportsPinning** 元素時，您的增益集內容不可以是固定的，且它必須清楚顯示信箱中已開啟或選定的郵件之相關資料。

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

如需完整的範例，請參閱 [command-demo 資訊清單](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml)內的 `msgReadOpenPaneButton` 控制項。

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>根據目前所選訊息來處理 UI 更新

若要更新工作窗格的 UI 或以目前項目為基礎的內部變數，您將需要註冊事件處理常式以獲得變更的通知。

### <a name="implement-the-event-handler"></a>實作事件處理常式

事件處理常式應該接受單一參數，也就是物件常值。此物件的 `type` 屬性將會設為 `Office.EventType.ItemChanged`。當呼叫事件時，`Office.context.mailbox.item` 物件已經更新以反映目前選取的項目。

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

### <a name="register-the-event-handler"></a>註冊事件處理常式

使用 [Office.context.mailbox.addHandlerAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#addHandlerAsync) 方法來註冊 `Office.EventType.ItemChanged` 事件的處理常式。應該在 `Office.initialize` 函式中為您的工作窗格完成此作業。

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="additional-resources"></a>其他資源

如需能實作可釘選工作窗格的增益集，請參閱 GitHub 上的 [command-demo](https://github.com/jasonjoh/command-demo) 範例。