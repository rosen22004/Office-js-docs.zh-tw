 

# <a name="mailboxenums"></a>MailboxEnums

## [Office](Office.md).MailboxEnums

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫|

### <a name="members"></a>成員

#### <a name="attachmenttype-:string"></a>AttachmentType︰字串

指定附件的類型。僅限撰寫模式。

AttachmentType

##### <a name="type:"></a>類型：

*   字串

##### <a name="properties:"></a>屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`File`| String|附件為檔案。|
|`Item`| String|附件為 Exchange 項目。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫|
#### <a name="entitytype-:string"></a>EntityType︰字串

指定實體的類型。僅限撰寫模式。

EntityType

##### <a name="type:"></a>類型：

*   字串

##### <a name="properties:"></a>屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`Address`| String|指定實體為郵寄地址。|
|`Contact`| 字串|指定實體為連絡人。|
|`EmailAddress`| 字串|指定實體為 SMTP 電子郵件地址。|
|`MeetingSuggestion`| String|指定實體為會議建議。|
|`PhoneNumber`| String|指定實體為美國電話號碼。|
|`TaskSuggestion`| 字串|指定實體為工作建議。|
|`URL`| String|指定實體為網際網路 URL。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫|
#### <a name="itemtype-:string"></a>ItemType：字串

指定項目的類型。僅限撰寫模式。

ItemType

##### <a name="type:"></a>類型：

*   字串

##### <a name="properties:"></a>屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`Message`| String|電子郵件、會議邀請，會議回覆或會議取消。|
|`Appoinment`| 字串|約會項目。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫|
#### <a name="recipienttype-:string"></a>RecipientType：字串

指定約會的收件者類型。僅限撰寫模式。

RecipientType

##### <a name="type:"></a>類型：

*   字串

##### <a name="properties:"></a>屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`Other`| String|收件者不是其中一個其他收件者類型。|
|`DistributionList`| String|收件者是包含電子郵件地址清單的通訊群組清單。|
|`User`| String|收件者是在 Exchange Server 上的 SMTP 電子郵件地址。|
|`ExternalUser`| String|收件者不是在 Exchange Server 上的 SMTP 電子郵件地址。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|適用的 Outlook 模式| 撰寫|
#### <a name="responsetype-:string"></a>ResponseType︰字串

指定會議邀請的回應類型。僅限撰寫模式。

ResponseType

##### <a name="type:"></a>類型：

*   字串

##### <a name="properties:"></a>屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`None`| String|沒有來自出席者的回應。|
|`Organizer`| 字串|出席者為會議召集人。|
|`Tentative`| 字串|出席者已暫訂接受會議邀請。|
|`Accepted`| String|出席者已接受會議邀請。|
|`Declined`| String|出席者已拒絕會議邀請。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫|
