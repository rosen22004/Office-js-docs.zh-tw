 

# MailboxEnums

## [Office](Office.md).MailboxEnums

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|

### 成員

#### AttachmentType︰字串

指定附件的類型。

AttachmentType

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 值 | 描述|
|---|---|---|---|
|`File`| String|`file`|附件為檔案。|
|`Item`| String|`item`|附件為 Exchange 項目。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|
#### EntityType︰字串

指定實體的類型。

EntityType

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 值 | 描述|
|---|---|---|---|
|`Address`| String|`address`|指定實體為郵寄地址。|
|`Contact`| 字串|`contact`|指定實體為連絡人。|
|`EmailAddress`| 字串|`emailAddress`|指定實體為 SMTP 電子郵件地址。|
|`MeetingSuggestion`| String|`meetingSuggestion`|指定實體為會議建議。|
|`PhoneNumber`| String|`phoneNumber`|指定實體為美國電話號碼。|
|`TaskSuggestion`| 字串|`taskSuggestion`|指定實體為工作建議。|
|`URL`| String|`url`|指定實體為網際網路 URL。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|
#### ItemNotificationMessageType︰字串

指定約會或郵件的通知訊息類型。

ItemNotificationMessageType

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 值 | 描述|
|---|---|---|---|
|`ProgressIndicator`| 字串|`progressIndicator`|NotificationMessage 為進度列指示器。|
|`InformationalMessage`| String|`informationalMessage`|NotificationMessage 為資訊訊息。|
|`ErrorMessage`| String|`errorMessage`|NotificationMessage 為錯誤訊息。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|適用的 Outlook 模式| 撰寫或讀取|
#### ItemType：字串

指定項目的類型。

ItemType

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 值 | 描述|
|---|---|---|---|
|`Message`| String|`message`|電子郵件、會議邀請，會議回覆或會議取消。|
|`Appointment`| 字串|`appointment`|約會項目。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|
#### RecipientType：字串

指定約會的收件者類型。

RecipientType

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 值 | 描述|
|---|---|---|---|
|`Other`| String|`other`|收件者不是其中一個其他收件者類型。|
|`DistributionList`| String|`distributionList`|收件者是包含電子郵件地址清單的通訊群組清單。|
|`User`| String|`user`|收件者是在 Exchange Server 上的 SMTP 電子郵件地址。|
|`ExternalUser`| String|`externalUser`|收件者不是在 Exchange Server 上的 SMTP 電子郵件地址。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|適用的 Outlook 模式| 撰寫或讀取|
#### ResponseType︰字串

指定會議邀請的回應類型。

ResponseType

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 值 | 描述|
|---|---|---|---|
|`None`| String|`none`|沒有來自出席者的回應。|
|`Organizer`| 字串|`organizer`|出席者為會議召集人。|
|`Tentative`| 字串|`tentative`|出席者已暫訂接受會議邀請。|
|`Accepted`| String|`accepted`|出席者已接受會議邀請。|
|`Declined`| String|`declined`|出席者已拒絕會議邀請。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|

#### RestVersion：字串

指定對應到 REST 格式的項目 ID 的 REST API 的版本。 

RestVersion

##### 類型：

*   字串

##### 屬性：

|名稱| 類型	| 值 | 描述|
|---|---|---|---|
|`v1_0`| 字串|`v1.0`|1.0 版。|
|`v2_0`| 字串|`v2.0`|2.0 版。|
|`Beta`| 字串|`beta`|Beta 版。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|適用的 Outlook 模式| 撰寫或讀取|
