# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>新增 Outlook Moblie 的增益集命令支援

> **注意**：目前只有 iOS 版 Outlook 支援 Outlook Moblie 的增益集命令。

使用 Outlook Mobile 中的增益集命令可讓使用者存取 Windows 版 Outlook、Mac 版 Outlook 及 Web 上的 Outlook 中已經具備的相同功能 (但有一些[限制](#code-considerations))。新增 Outlook Mobile 支援時，需要更新增益集資訊清單並可能需要變更行動裝置的程式碼。

## <a name="updating-the-manifest"></a>更新資訊清單

在 Outlook Moblie 中啟用增益集命令的第一個步驟是在增益集資訊清單中定義這些命令。**VersionOverrides** v1.1 版結構描述會為行動裝置定義新的外觀尺寸，即 [MobileFormFactor](../../reference/manifest/mobileformfactor.md)。

這個元素包含在行動用戶端中載入增益集的所有資訊。這可讓您針對行動體驗定義完全不同的 UI 元素和 JavaScript 檔案。

下列範例顯示 **MobileFormFactor** 元素中的單一工作窗格按鈕。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Control xsi:type="MobileButton" id="TaskPane1Btn">
        <Label resid="residTaskPaneButton0Name" />
        <Icon xsi:type="bt:MobileIconList">
          <bt:Image size="25" scale="1" resid="tp0icon" />
          <bt:Image size="25" scale="2" resid="tp0icon" />
          <bt:Image size="25" scale="3" resid="tp0icon" />

          <bt:Image size="32" scale="1" resid="tp0icon" />
          <bt:Image size="32" scale="2" resid="tp0icon" />
          <bt:Image size="32" scale="3" resid="tp0icon" />

          <bt:Image size="48" scale="1" resid="tp0icon" />
          <bt:Image size="48" scale="2" resid="tp0icon" />
          <bt:Image size="48" scale="3" resid="tp0icon" />
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="residTaskpaneUrl" />
        </Action>
      </Control>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

這非常類似於 [DesktopFormFactor](../../reference/manifest/desktopformfactor.md) 元素中顯示的元素，但有一些顯著的差異。

- 不會使用 [OfficeTab](../../reference/manifest/officetab.md) 元素。
- [ExtensionPoint](../../reference/manifest/exensionpoint.md) 元素必須只有一個子元素。如果增益集只新增一個按鈕，則子元素應該是 [Control](../../reference/manifest/control.md) 元素。如果增益集加入一個以上的按鈕，則子元素應該是 [Group](../../reference/manifest/group.md) 元素，其中包含多個 `Control` 元素。
- 有沒等同於 `Control` 元素的 `Menu` 類型。
- 不會使用 [Supertip](../../reference/manifest/supertip.md) 元素。
- 所需的圖示大小有所不同。Mobile 增益集至少必須支援 25 x 25、32x32 和 48x48 像素的圖示。

## <a name="code-considerations"></a>程式碼考量

設計適用於行動裝置的增益集會有一些其他考量。

### <a name="use-rest-instead-of-exchange-web-services"></a>使用 REST 而非 Exchange Web 服務

Outlook Mobile 不支援 [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法。可能的話，增益集應該會偏好從 Office.js API 取得資訊。如果增益集需要 Office.js API 並未公開的資訊，那麼就應該使用 [Outlook REST APIs](https://dev.outlook.com/restapi/reference) 來存取使用者的信箱。 

信箱需求集合 1.5 引入了新的 [Office.context.mailbox.getCallbackTokenAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#getCallbackTokenAsync) 版本 (其可要求與 REST API 相容的存取權杖)，以及新的 [Office.context.mailbox.restUrl](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#restUrl) 屬性 (其可用來尋找使用者的 REST API 端點)。

### <a name="pinch-zoom"></a>捏合縮放

根據預設，使用者可以使用「捏合縮放」手勢來放大工作窗格。如果這對於您的案例毫無意義，請務必在 HTML 中停用捏合縮放。

### <a name="closing-taskpanes"></a>關閉工作窗格

在 Outlook Mobile 中，工作窗格會佔用整個螢幕，且預設會要求使用者關閉這些工作窗格，才能返回郵件。當您的案例完成時，請考慮使用 [Office.context.ui.closeContainer](https://dev.outlook.com/reference/add-ins/1.5/Office.context.ui.html#closeContainer) 方法來關閉工作窗格。

### <a name="compose-mode-and-appointments"></a>撰寫模式和約會

Outlook Mobile 中的增益集目前只支援在讀取郵件啟動。在撰寫郵件時或在檢視或撰寫約會時，不會啟動增益集。

### <a name="unsupported-apis"></a>不支援的 API

Outlook Mobile 不支援下列 API。

  - [Office.context.officeTheme](../../reference/outlook/Office.context.md)
  - [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToEwsId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToRestId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.resources](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyAllForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntities](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatches](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)