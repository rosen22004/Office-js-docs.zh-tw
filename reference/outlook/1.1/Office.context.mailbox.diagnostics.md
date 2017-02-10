

# <a name="diagnostics"></a>診斷

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

提供診斷資訊給 Outlook 增益集。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="members"></a>成員

####  <a name="hostname-string"></a>hostName：字串

取得代表主機應用程式名稱的字串。

字串可能是下列其中一個值：`Outlook`、`Mac Outlook`、`OutlookIOS` 或 `OutlookWebApp`。

##### <a name="type"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
####  <a name="hostversion-string"></a>hostVersion：字串

取得代表主機應用程式或 Exchange Server的版本的字串。

如果郵件增益集在 Outlook 桌面用戶端或 iOS 版 Outlook 上執行，`hostVersion` 屬性會傳回主機應用程式 - Outlook 的版本。在 Outlook Web App 中，該屬性會傳回 Exchange Server 的版本。`15.0.468.0` 字串即是一例。

##### <a name="type"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
####  <a name="owaview-string"></a>OWAView：字串

取得字串代表 Outlook Web App 目前檢視的字串。

傳回的字串可能是下列其中一個值：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。

如果主機應用程式不是 Outlook Web App，存取這個屬性會導致 `undefined`。

Outlook Web App 具有三個與螢幕和視窗寬度，以及可顯示之資料行數目相對應的檢視︰

*   `OneColumn`：會在窄螢幕時顯示。Outlook Web App 會將這個單欄式配置用在整個智慧型手機的螢幕上。
*   `TwoColumns`：會在螢幕較寬時顯示。Outlook Web App 會將這個檢視用在大部分的平板電腦上。
*   `ThreeColumns`：會在寬螢幕時顯示。例如，Outlook Web App 會將這個檢視用在桌面電腦上的全螢幕視窗中。

##### <a name="type"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
