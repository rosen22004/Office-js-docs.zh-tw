

# <a name="item"></a>項目

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). item

`item` 命名空間用來存取目前所選的郵件、會議邀請或約會。您可以使用 [itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype) 屬性來判斷 `item` 的類型。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="example"></a>範例

下列的 JavaScript 程式碼範例示範如何在 Outlook 中存取目前項目的 `subject` 屬性。

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>成員

#### <a name="attachments-:array.<[attachmentdetails](simple-types.md#attachmentdetails)>"></a>附件：Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

取得項目附件的陣列。僅限閱讀模式。

##### <a name="type:"></a>類型：

*   Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

下列程式碼會建置 HTML 字串，內含目前項目上所有附件的詳細資料。

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-:[recipients](recipients.md)"></a>密件副本︰[收件者](Recipients.md)

取得或設定郵件 [密件副本] 行上的收件者。僅限撰寫模式。

##### <a name="type:"></a>類型：

*   [收件者](Recipients.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="example"></a>範例

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-:[body](body.md)"></a>本文︰[本文](Body.md)

取得提供方法來管理項目本文的物件。

##### <a name="type:"></a>類型：

*   [內文](Body.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
####  副本：Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[收件者](Recipients.md)

取得或設定郵件的副本收件者。

##### <a name="read-mode"></a>閱讀模式

`cc` 屬性傳回陣列，包含郵件 [副本] 列上所列出每個收件者的 `EmailAddressDetails` 物件。這個集合限制最多為 100 名成員。

##### <a name="compose-mode"></a>撰寫模式

`cc` 屬性傳回 `Recipients` 物件，提供方法來管理郵件 [副本] 列上的收件者。

##### <a name="type:"></a>類型：

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [收件者](Recipients.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="(nullable)-conversationid-:string"></a>(可為 null) conversationId：字串

取得包含特定郵件的電子郵件交談識別碼。

如果郵件應用程式是在讀取模式中啟動，或是在撰寫模式中回應；您便可以取得這個屬性的整數。如果使用者隨後在傳送回覆時，變更回覆郵件的主旨，那麼該郵件的交談識別碼將會變更，而且您稍早所取得的值將無法再套用。

您在撰寫模式中取得新項目的此屬性為 null。如果使用者設定主旨並儲存項目，`conversationId` 屬性會傳回一個值。

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
#### <a name="datetimecreated-:date"></a>dateTimeCreated︰日期

取得項目已建立的時間與日期。僅限閱讀模式。

##### <a name="type:"></a>類型：

*   日期

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-:date"></a>dateTimeModified︰日期

取得項目上次修改的日期與時間。僅限閱讀模式。

##### <a name="type:"></a>類型：

*   日期

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-:date|[time](time.md)"></a>結束：日期|[時間](Time.md)

取得或設定約會要結束的日期和時間。

`end` 屬性以國際標準時間 (UTC) 中的日期和時間值來表示。您可以使用 [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法，將結束屬性值轉換成用戶端的本機日期和時間。

##### <a name="read-mode"></a>閱讀模式

`end` 屬性傳回 `Date` 物件。

##### <a name="compose-mode"></a>撰寫模式

`end` 屬性傳回 `Time` 物件。

使用 [`Time.setAsync`](Time.md#setasync) 方法來設定結束時間時，您應該使用 [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) 方法，將用戶端上的本機時間轉換成伺服器的 UTC。

##### <a name="type:"></a>類型：

*   日期 | [時間](Time.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

下列範例使用 `Time` 物件的 [`setAsync`](Time.md#setasync) 方法，來設定撰寫模式中約會的結束時間。

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-:[emailaddressdetails](simple-types.md#emailaddressdetails)"></a>寄件者︰[EmailAddressDetails](simple-types.md#emailaddressdetails)

取得郵件寄件者的電子郵件地址。僅限閱讀模式。

`from` 和 [`sender`](Office.context.mailbox.item.md#sender) 屬性代表同一個人，除非郵件是由代理人所傳送。在此情況下，`from` 屬性表示委派人，而寄件者屬性即表示代理人。

##### <a name="type:"></a>類型：

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
#### <a name="internetmessageid-:string"></a>internetMessageId：字串

取得電子郵件訊息的網際網路郵件識別碼。僅限閱讀模式。

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-:string"></a>itemClass：字串

取得選取項目的 Exchange Web 服務項目類別。僅限閱讀模式。

`itemClass` 屬性會指定選取項目的郵件類別。下列是郵件或約會項目預設的郵件類別。

| 類型 | 描述 | 項目類別 |
| --- | --- | --- |
| 約會項目 | 這些是項目類別 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的行事曆項目。 | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| 郵件項目 | 這些包括有預設郵件類別 `IPM.Note` 的電子郵件訊息和使用 `IPM.Schedule.Meeting` 作為基礎郵件類別的會議邀請、回覆和取消。 | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

您可以建立可擴充預設郵件類別的自訂郵件類別，例如，自訂約會郵件類別 `IPM.Appointment.Contoso`。

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="(nullable)-itemid-:string"></a>(可為 null) itemId：字串

取得目前項目的 Exchange Web 服務項目識別碼。僅限閱讀模式。

由 `itemId` 屬性所傳回的識別碼和 Exchange Web 服務項目的識別碼相同。`itemId` 屬性不同於 Outlook 項目識別碼。

對於未儲存至儲存區的項目，`itemId` 屬性會在撰寫模式下傳回 `null`。如果項目識別碼為必要，[`saveAsync`](Office.context.mailbox.item.md#saveAsync) 方法可以用來將項目儲存至儲存區，其會將回呼函數中 [`AsyncResult.value`](simple-types.md#asyncresult) 參數內的項目識別碼傳回。

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

下列程式碼會檢查項目識別碼是否存在。如果 `itemId` 屬性傳回 `null` 或 `undefined`，其會將項目儲存至儲存區，並自非同步結果取得項目識別碼。

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-:[office.mailboxenums.itemtype](office.mailboxenums.md#itemtype-string)"></a>itemType :[Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

取得執行個體所表示的項目類型。

`itemType` 屬性傳回其中一個 `ItemType` 列舉值，指出 `item` 物件執行個體是否為郵件或約會。

##### <a name="type:"></a>類型：

*   [Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-:string|[location](location.md)"></a>地點：字串|[地點](Location.md)

取得或設定約會的地點。

##### <a name="read-mode"></a>閱讀模式

`location` 屬性傳回包含約會地點的字串。

##### <a name="compose-mode"></a>撰寫模式

`location` 屬性傳回 `Location` 物件，提供方法用來取得和設定約會地點。

##### <a name="type:"></a>類型：

*   字串 | [地點](Location.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-:string"></a>normalizedSubject︰字串

取得移除所有前置詞 (包括 `RE:` 和 `FWD:`) 的項目主旨。僅限閱讀模式。

NormalizedSubject 屬性取得項目主旨，內含由電子郵件程式新增的任一標準前置詞 (例如 `RE:` 和 `FW:`)。若要取得項目主旨，而前置詞維持不變，請使用 [`subject`](Office.context.mailbox.item.md#subject) 屬性。

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-:array.<[emailaddressdetails](simple-types.md#emailaddressdetails)>|[recipients](recipients.md)"></a>optionalAttendees：Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[收件者](Recipients.md)

取得或設定列席者的電子郵件地址清單。

##### <a name="read-mode"></a>閱讀模式

`optionalAttendees` 屬性傳回陣列，包含會議每個列席者的 `EmailAddressDetails` 物件。

##### <a name="compose-mode"></a>撰寫模式

`optionalAttendees` 屬性傳回 `Recipients` 物件，提供方法來取得及設定會議列席者。

##### <a name="type:"></a>類型：

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [收件者](Recipients.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-:[emailaddressdetails](simple-types.md#emailaddressdetails)"></a>召集人︰[EmailAddressDetails](simple-types.md#emailaddressdetails)

取得指定會議的會議召集人電子郵件地址。僅限閱讀模式。

##### <a name="type:"></a>類型：

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-:array.<[emailaddressdetails](simple-types.md#emailaddressdetails)>|[recipients](recipients.md)"></a>requiredAttendees：Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[收件者](Recipients.md)

取得或設定出席者的電子郵件地址清單。

##### <a name="read-mode"></a>閱讀模式

`requiredAttendees` 屬性傳回陣列，包含會議每個出席者的 `EmailAddressDetails` 物件。

##### <a name="compose-mode"></a>撰寫模式

`requiredAttendees` 屬性傳回 `Recipients` 物件，提供方法來取得及設定會議出席者。

##### <a name="type:"></a>類型：

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [收件者](Recipients.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="resources-:[emailaddressdetails](simple-types.md#emailaddressdetails)"></a>資源：[EmailAddressDetails](simple-types.md#emailaddressdetails)

取得約會所需的資源。僅限閱讀模式。

##### <a name="type:"></a>類型：

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
#### <a name="sender-:[emailaddressdetails](simple-types.md#emailaddressdetails)"></a>寄件者︰[EmailAddressDetails](simple-types.md#emailaddressdetails)

取得電子郵件訊息的寄件者電子郵件地址。僅限閱讀模式。

[`from`](Office.context.mailbox.item.md#from) 和 `sender` 屬性代表同一個人，除非郵件是由代理人所傳送。在此情況下，`from` 屬性表示委派人，而寄件者屬性即表示代理人。

##### <a name="type:"></a>類型：

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-:date|[time](time.md)"></a>開始：日期|[時間](Time.md)

取得或設定約會要開始的日期和時間。

`start` 屬性以國際標準時間 (UTC) 中的日期和時間值來表示。您可以使用 [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法，將值轉換成用戶端的本機日期和時間。

##### <a name="read-mode"></a>閱讀模式

`start` 屬性傳回 `Date` 物件。

##### <a name="compose-mode"></a>撰寫模式

`start` 屬性傳回 `Time` 物件。

使用 [`Time.setAsync`](Time.md#setasync) 方法來設定開始時間時，您應該使用 [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) 方法，將用戶端上的本機時間轉換成伺服器的 UTC。

##### <a name="type:"></a>類型：

*   日期 | [時間](Time.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

下列範例使用 `Time` 物件的 [`setAsync`](Time.md#setasync) 方法，來設定撰寫模式中約會的開始時間。

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-:string|[subject](subject.md)"></a>主旨︰字串|[主旨](Subject.md)

取得或設定項目的 [主旨] 欄位中出現的描述。

`subject` 屬性取得或設定電子郵件伺服器所傳送之項目的完整主旨。

##### <a name="read-mode"></a>閱讀模式

`subject` 屬性傳回字串。使用 [`normalizedSubject`](Office.context.mailbox.item.md#normalizedsubject-string) 屬性，以取得減去任何前置字元前置詞的主旨，例如 `RE:` 和 `FW:`。

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>撰寫模式

`subject` 屬性傳回 `Subject` 物件，提供方法來取得及設定主旨。

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type:"></a>類型：

*   字串 | [主旨](Subject.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
####  收件者：Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[收件者](Recipients.md)

取得或設定電子郵件訊息的收件者。

##### <a name="read-mode"></a>閱讀模式

`to` 屬性傳回陣列，包含郵件 [收件者] 列上所列出每個收件者的 `EmailAddressDetails` 物件。這個集合限制最多為 100 名成員。

##### <a name="compose-mode"></a>撰寫模式

`to` 屬性傳回 `Recipients` 物件，提供方法來管理郵件 [收件者] 列上的收件者。

##### <a name="type:"></a>類型：

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [收件者](Recipients.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>方法

####  <a name="addfileattachmentasync(uri,-attachmentname,-[options],-[callback])"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

將檔案新增至郵件或約會做為附件。

`addFileAttachmentAsync` 方法將檔案上傳至指定的 URI，並在撰寫模式中將它附加到項目。

您可以隨後以 [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) 方法使用識別碼，以便在相同的工作階段中移除附件。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`uri`| string||提供要附加至郵件或約會的檔案位置 URI。最大長度為 2048 個字元。|
|`attachmentName`| string||正在上傳附件時，會顯示附件名稱。最大長度為 255 個字元。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>一旦成功，附件識別碼會在 `asyncResult.value` 屬性中提供。<br/>如果上載附件失敗，`asyncResult` 物件將包含 `Error` 物件，提供錯誤的描述。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>說明</th></tr></thead><tbody><tr><td><code>AttachmentSizeExceeded</code></td><td>附件大於允許大小。</td></tr><tr><td><code>FileTypeNotSupported</code></td><td>附件具有不允許的副檔名。</td></tr><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>郵件或約會有太多的附件。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="example"></a>範例

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasync(itemid,-attachmentname,-[options],-[callback])"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

將 Exchange 項目，例如訊息，新增至郵件或約會做為附件。

`addItemAttachmentAsync` 方法會在撰寫模式中將具有指定 Exchange 識別碼的項目附加至項目。如果您指定回呼方法，方法是以一個參數 `asyncResult` 呼叫，其包含附件識別碼或是在附加項目時指出所發生之錯誤的程式碼。如有必要的話，您可以使用 `options` 參數，將狀態資訊傳遞至回呼方法。

您可以隨後以 [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback) 方法使用識別碼，以便在相同的工作階段中移除附件。

如果您的 Office 增益集正在 Outlook Web App 中執行，`addItemAttachmentAsync` 方法可將項目附加至您正在編輯項目以外的項目；不過，不支援此方法，也不建議這麼做。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`itemId`| string||要附加的項目 Exchange 識別碼。最大長度為 100 個字元。|
|`attachmentName`| string||要附加的項目主旨。最大長度為 255 個字元。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>一旦成功，附件識別碼會在 `asyncResult.value` 屬性中提供。<br/>如果新增附件失敗，`asyncResult` 物件將包含 `Error` 物件，提供錯誤的描述。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>說明</th></tr></thead><tbody><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>郵件或約會有太多的附件。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="example"></a>範例

下列範例會以 `My Attachment` 的名稱將現有 Outlook 項目新增為附件。

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallform(formdata)"></a>displayReplyAllForm(formData)

顯示包含所選郵件的寄件者和所有收件者或召集人，以及所選約會的所有出席者的回覆表單。

在 Outlook Web App 中，回覆表單會顯示為 3 欄式檢視中的彈出式表單，以及在 2 欄或 1 欄式檢視中的快顯表單。

如果任何字串參數超過限制，`displayReplyAllForm` 會拋出例外狀況。

> **附註：**在對 `displayReplyAllForm` 的呼叫中包含附件的能力，在需求集 1.1 中不支援。附件支援已新增至 `displayReplyAllForm` 需求集 1.2 及以上。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 說明|
|---|---|---|
|`formData`| string &#124; 物件|包含文字和 HTML，且代表回覆表單本文的字串。字串限制為 32 KB。<br/>**或**<br/>包含本文或附件資料和回呼函數的物件。物件定義如下：<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>字串</td><td>&lt;選用&gt;</td><td>包含文字和 HTML，且代表回覆表單本文的字串。字串限制為 32 KB。</td></tr><tr><td><code>callback</code></td><td>函數</td><td>&lt;選用&gt;</td><td>當方法完成時，在 <code>callback</code> 參數中傳遞的函數會以單一參數 <code>asyncResult</code>，也就是 <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a> 物件進行呼叫。如需詳細資訊，請參閱<a href="tutorial-asynchronous.html">使用非同步方法</a>。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="examples"></a>範例

下列程式碼會將字串傳遞至 `displayReplyAllForm` 函數。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

以空的本文回覆。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

只以本文回覆。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

以本文及回呼回覆。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyform(formdata)"></a>displayReplyForm(formData)

顯示只包含所選郵件的寄件者或所選約會召集人的回覆表單。

在 Outlook Web App 中，回覆表單會顯示為 3 欄式檢視中的彈出式表單，以及在 2 欄或 1 欄式檢視中的快顯表單。

如果任何字串參數超過限制，`displayReplyForm` 會拋出例外狀況。

> **附註：**在對 `displayReplyForm` 的呼叫中包含附件的能力，在需求集 1.1 中不支援。附件支援已新增至 `displayReplyForm` 需求集 1.2 及以上。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 說明|
|---|---|---|
|`formData`| string &#124; 物件|包含文字和 HTML，且代表回覆表單本文的字串。字串限制為 32 KB。<br/>**或**<br/>包含本文或附件資料和回呼函數的物件。物件定義如下：<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>字串</td><td>&lt;選用&gt;</td><td>包含文字和 HTML，且代表回覆表單本文的字串。字串限制為 32 KB。</td></tr><tr><td><code>callback</code></td><td>函數</td><td>&lt;選用&gt;</td><td>當方法完成時，在 <code>callback</code> 參數中傳遞的函數會以單一參數 <code>asyncResult</code>，也就是 <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a> 物件進行呼叫。如需詳細資訊，請參閱<a href="tutorial-asynchronous.html">使用非同步方法</a>。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="examples"></a>範例

下列程式碼會將字串傳遞至 `displayReplyForm` 函數。

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

以空的本文回覆。

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

只以本文回覆。

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

以本文及回呼回覆。

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities()-→-{[entities](simple-types.md#entities)}"></a>getEntities() → {[實體](simple-types.md#entities)}

取得在選取項目中所找到的實體。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="returns:"></a>傳回：

類型：
[實體](simple-types.md#entities)

##### <a name="example"></a>範例

下列範例會在目前項目上存取連絡人實體。

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytype(entitytype)-→-(nullable)-{array.<(string|[contact](simple-types.md#contact)|[meetingsuggestion](simple-types.md#meetingsuggestion)|[phonenumber](simple-types.md#phonenumber)|[tasksuggestion](simple-types.md#tasksuggestion))>}"></a>getEntitiesByType(entityType) → (可為 null) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

取得指定實體類型 (在選取項目中所找到) 的所有實體陣列。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#.entitytype-string)|其中一個 EntityType 列舉值。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 讀取|

##### <a name="returns:"></a>傳回：

如果在 `entityType` 中傳遞的值不是 `EntityType` 列舉的有效成員，則方法會傳回 null。如果沒有指定類型的實體存在於項目上，則方法會傳回空陣列。否則，在傳回陣列中的物件類型會視 `entityType` 參數中所要求的實體類型而定。

當使用此方法的最低權限等級為 **限制**，一些實體類型需要 **ReadItem** 才能存取，如下表中所指定。

| `entityType` 的值 | 傳回陣列中的物件類型 | 必要的權限等級 |
| --- | --- | --- |
| `Address` | String | **Restricted** |
| `Contact` | 連絡人 | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restricted** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **Restricted** |

類型：Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))></dd>


##### <a name="example"></a>範例

下列範例會示範如何存取代表目前項目主旨或本文中郵寄地址的字串陣列。

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbyname(name)-→-(nullable)-{array.<(string|[contact](simple-types.md#contact)|[meetingsuggestion](simple-types.md#meetingsuggestion)|[phonenumber](simple-types.md#phonenumber)|[tasksuggestion](simple-types.md#tasksuggestion))>}"></a>getFilteredEntitiesByName(name) → (可為 null) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

在選取項目中傳回已知實體，該項目會傳遞在資訊清單 XML 檔案中定義的命名篩選。

  `getFilteredEntitiesByName` 方法傳回符合規則運算式的實體，該運算式是在資訊清單 XML 檔案的 [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/office/fp161166.aspx) 規則項目中所定義，該規則元素具有指定的 `FilterName` 元素值。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`name`| string|定義要符合篩選的 `ItemHasKnownEntity` 規則元素名稱。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="returns:"></a>傳回：

如果資訊清單中沒有任何 `ItemHasKnownEntity` 項目是具有符合 `name` 參數的 `FilterName` 項目值，則方法會傳回 `null`。如果 `name` 參數不符合資訊清單中的 `ItemHasKnownEntity` 項目，但在目前相符的項目中沒有實體，則方法會傳回空陣列。


類型：Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>


#### <a name="getregexmatches()-→-{object}"></a>getRegExMatches() → {物件}

在選取項目中傳回符合規則運算式的字串值，該值是在資訊清單 XML 檔中所定義。

`getRegExMatches` 方法傳回符合規則運算式的字串，該運算式是在資訊清單 XML 檔案的每個 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 規則項目中所定義。對於 `ItemHasRegularExpressionMatch` 規則，相符的字串必須出現在由該規則所指定之項目的屬性中。`PropertyName` 簡單類型定義所支援的屬性。

例如，假設增益集資訊清單有下列 `Rule` 項目︰

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

從 `getRegExMatches` 傳回的物件會有兩個屬性︰`fruits` 和 `veggies`。

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

如果您在項目的本文屬性上指定 `ItemHasRegularExpressionMatch` 規則，規則運算式應該進一步篩選本文，且不應該嘗試傳回項目的整個本文。使用規則運算式，例如 `.*` 來取得項目的整個本文，不會永遠都傳回預期的結果。相反地，請使用 [`Body.getAsync`](Body.md#getAsync) 方法，以擷取整個本文。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="returns:"></a>傳回：

物件包含符合規則運算式的字串陣列，該運算式是在資訊清單 XML 檔案中所定義。每個陣列名稱等於相符 `ItemHasRegularExpressionMatch` 規則 `RegExName` 屬性或相符 `ItemHasKnownEntity` 規則 `FilterName` 屬性的相對應值。

<dl class="param-type">

<dt>類型</dt>

<dd>物件</dd>

</dl>

##### <a name="example"></a>範例

下列範例示範如何存取規則運算式相符的陣列 <rule>元素 `fruits` 和 `veggies`，其是在資訊清單中指定。</rule>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbyname(name)-→-(nullable)-{array.<string>}"></a>getRegExMatchesByName(name) → (可為 null) {Array.<String>}

在選取項目中傳回符合命名規則運算式的字串值，該值是在資訊清單 XML 檔中所定義。

`getRegExMatchesByName` 方法傳回符合規則運算式的字串，該運算式是在資訊清單 XML 檔案的 `ItemHasRegularExpressionMatch` 規則項目中所定義，該規則項目具有指定的 `RegExName` 項目值。

如果您在項目的本文屬性上指定 `ItemHasRegularExpressionMatch` 規則，規則運算式應該進一步篩選本文，且不應該嘗試傳回項目的整個本文。使用規則運算式，例如 `.*` 來取得項目的整個本文，不會永遠都傳回預期的結果。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`name`| string|定義要符合篩選的 `ItemHasRegularExpressionMatch` 規則項目名稱。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="returns:"></a>傳回：

陣列包含符合規則運算式的字串，該運算式是在資訊清單 XML 檔案內所定義。

<dl class="param-type">

<dt>類型</dt>

<dd>陣列。<String></dd>

</dl>

##### <a name="example"></a>範例

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasync(coerciontype,-[options],-callback)-→-{string}"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

以非同步方式從郵件主旨或本文傳回選取資料。

如果沒有選取範圍，但游標位於本文或主旨中，方法會傳回選取資料為 null。如果選取本文或主旨以外的欄位，則這個方法會傳回 `InvalidSelection` 錯誤。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||要求資料的格式。如果是文字，這個方法會傳回純文字當做字串，移除任何出現的 HTML 標籤。如果是 HTML，這個方法會傳回選取的文字，不論是純文字或 HTML。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

若要從回呼方法存取選取的資料，請呼叫 `asyncResult.value.data`。若要存取來自選取範圍的來源屬性，請呼叫 `asyncResult.value.sourceProperty`，這會是 `body` 或 `subject`。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="returns:"></a>傳回：

選取的資料當做字串，是由 `coercionType` 決定格式。

<dl class="param-type">

<dt>類型</dt>

<dd>字串</dd>

</dl>

##### <a name="example"></a>範例

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasync(callback,-[usercontext])"></a>loadCustomPropertiesAsync(callback, [userContext])

以非同步方式載入選取項目上此增益集的自訂屬性。

以每個應用程式和每個項目為基礎，將自訂屬性儲存為索引鍵/值組。這個方法會在回呼中傳回 `CustomProperties` 物件，此物件提供方法來存取目前項目和目前增益集特有的自訂屬性。自訂屬性尚未對項目進行加密，所以這不應該用來做為安全儲存。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

提供自訂屬性，做為在 `asyncResult.value` 屬性中的 [`CustomProperties`](CustomProperties.md) 物件。這個物件可用來取得、設定和移除來自項目的自訂屬性，並將變更儲存回伺服器的自訂屬性集。| 
|`userContext`| 物件| &lt;選用&gt;|開發人員可以提供任何他們想要在回呼函數中存取的物件。這個物件可由回呼函數中的 `asyncResult.asyncContext` 屬性來進行存取。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

下列程式碼範例示範如何使用 `loadCustomPropertiesAsync` 方法以非同步方式載入目前項目特有的自訂屬性。範例中也示範如何使用 `CustomProperties.saveAsync` 方法，將這些屬性儲存回伺服器。載入自訂屬性之後，程式碼範例會使用 `CustomProperties.get` 方法讀取自訂屬性 `myProp`，使用 `CustomProperties.set` 方法撰寫自訂屬性 `otherProp`，然後最後呼叫 `saveAsync` 方法，來儲存自訂屬性。

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasync(attachmentid,-[options],-[callback])"></a>removeAttachmentAsync(attachmentId, [options], [callback])

移除來自郵件或約會的附件。

`removeAttachmentAsync` 方法會將具有指定識別碼的附件自項目中移除。最佳作法是，唯有當相同郵件應用程式已在相同的工作階段中新增該附件時，您才應該使用附件識別碼來移除附件。在 Outlook Web App 和 OWA for Devices 中，附件識別碼只有在相同工作階段內才會有效。當使用者關閉應用程式時，工作階段會結束，或如果使用者開始在內嵌表單進行撰寫，接下來會跳出內嵌表單以便在個別視窗中繼續。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`attachmentId`| string||要移除的附件識別碼。字串的最大長度為 100 個字元。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>如果附件移除失敗，`asyncResult.error` 屬性將會包含錯誤碼與失敗原因。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>說明</th></tr></thead><tbody><tr><td><code>InvalidAttachmentId</code></td><td>附件識別碼不存在。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="example"></a>範例

下列程式碼移除識別碼為 '0' 的附件。

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```
