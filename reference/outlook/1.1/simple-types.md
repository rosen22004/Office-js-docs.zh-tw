

# 簡單類型

####  AsyncResult

會封裝非同步要求結果的物件，包括狀態及錯誤資訊 (如果要求失敗)。

##### 屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`asyncContext`| 物件|取得傳遞給叫用方法之選擇性 `asyncContext` 參數的物件，並保留傳遞時的狀態。|
|`error`| 錯誤|如果發生任何錯誤，取得提供錯誤描述的 Error 物件。|
|`status`| [Office.AsyncResultStatus](Office.md#.asyncresultstatus-string)|取得非同步作業的狀態。|
|`value`| 物件|取得這個非同步作業的裝載或內容 (如果有的話)。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|
#### AttachmentDetails

代表來自伺服器之項目上的附件。僅限閱讀模式。

傳回的 `AttachmentDetail` 物件陣列可當做 `attachments` 或 `Appointment` 物件的 `Message` 屬性。

##### 屬性：

|名稱| 類型	| 說明|
|---|---|---|
|`attachmentType`| [Office.MailboxEnums.AttachmentType](Office.MailboxEnums.md#attachmenttype-string)|取得指出附件類型的值。|
|`contentType`| 字串|取得附件的 MIME 內容類型。|
|`id`| String|取得附件的 Exchange 附件識別碼。|
|`isInline`| Boolean|取得指出附件是否應顯示在項目本文的值。|
|`name`| 字串|取得附件的名稱。|
|`size`| 數字|取得附件的大小 (以位元組為單位)。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
#### 連絡人

代表儲存在伺服器上的連絡人。僅限閱讀模式。

以 [`contacts`](simple-types.md#entities) 物件的 `Entities` 屬性傳回與電子郵件訊息或約會相關的連絡人清單，該物件是以使用中項目的 `getEntities` 或 `getEntitiesByType` 方法傳回。

##### 屬性：

|名稱| 類型	| 屬性| 說明|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;可為 null&gt;|包含與連絡人相關之郵寄和街道地址的字串陣列。|
|`businessName`| String| &lt;可為 null&gt;|包含與連絡人相關之公司名稱的字串。|
|`emailAddresses`| Array.&lt;String&gt;| &lt;可為 null&gt;|包含與連絡人相關之 SMTP 電子郵件地址的字串陣列。|
|`personName`| String| &lt;可為 null&gt;|包含與連絡人相關之人員姓名的字串。|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;可為 null&gt;|每個與連絡人相關之電話號碼均有一個 `PhoneNumber` 物件的陣列。|
|`urls`| Array.&lt;String&gt;| &lt;可為 null&gt;|包含與連絡人相關之網際網路 URL 的字串陣列。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 讀取|
####  EmailAddressDetails

提供電子郵件訊息或約會之寄件者或指定收件者的電子郵件屬性。

##### 類型：

*   物件

##### 屬性：

|名稱| 類型	| 說明|
|---|---|---|
|`appointmentResponse`| [Office.MailboxEnums.ResponseType](Office.MailboxEnums.md#responsetype-string)|取得出席者針對約會傳回的回應。這個屬性只適用於約會的出席者，以 [`optionalAttendees`](Office.context.mailbox.item.md#optionalattendees-arrayemailaddressdetailsrecipients) 或 [`requiredAttendees`](Office.context.mailbox.item.md#requiredattendees-arrayemailaddressdetailsrecipients) 屬性表示。這個屬性在其他案例中會傳回 `undefined`。|
|`displayName`| 字串|取得與電子郵件地址相關的顯示名稱。|
|`emailAddress`| 字串|取得 SMTP 電子郵件地址。|
|`recipientType`| [Office.MailboxEnums.RecipientType](Office.MailboxEnums.md#recipienttype-string)|取得收件者的電子郵件地址類型。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
#### EmailUser

代表 Exchange Server 上的電子郵件帳戶。

##### 屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`displayName`| String|取得與電子郵件地址相關的顯示名稱。|
|`emailAddress`| 字串|取得 SMTP 電子郵件地址。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
#### 實體

代表電子郵件訊息或約會中的實體集合。僅限閱讀模式。

當項目 (電子郵件訊息或約會) 含有一或多個伺服器已發現的實體時，`Entities` 物件是 `getEntities` 和 `getEntitiesByType` 方法傳回之實體陣列的容器。您可以在程式碼中使用這些實體，以將其他內容資訊提供給檢視者 (如前往項目中之地址的地圖)，或開啟撥號程式以撥打項目中的電話號碼。

如果項目中沒有屬性指定的實體類型，與該實體相關的屬性會是 `null`。例如，如果訊息包含街道地址和電話號碼，`addresses` 屬性和 `phoneNumbers` 屬性將含有這些資訊，而其他屬性則會是 `null`。

若要辨識為地址，字串必須包含至少有街道號碼、街道名稱、城市和郵遞區號等元素之子集的美國郵寄地址。

若要辨識為電話號碼，字串必須包含北美電話號碼格式。

實體辨識仰賴以蒐羅大量資料之機器學習為基礎的自然語言辨識。實體辨識不具決定性，成功有時需依靠項目中的特定內容。

當 `getEntitiesByType` 方法傳回屬性陣列時，只有指定之實體的屬性包含資料，其他所有屬性均會是 `null`。

##### 屬性：

|名稱| 類型	| 屬性| 說明|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;可為 null&gt;|取得電子郵件訊息或約會中的實體位址 (街道或郵寄地址)。|
|`contacts`| Array.&lt;[Contact](simple-types.md#contact)&gt;| &lt;可為 null&gt;|取得電子郵件地址或約會中的連絡人。|
|`emailAddresses`| Array.&lt;String&gt;| &lt;可為 null&gt;|取得電子郵件訊息或約會中的電子郵件地址。|
|`meetingSuggestions`| Array.&lt;[MeetingSuggestion](simple-types.md#meetingsuggestion)&gt;| &lt;可為 null&gt;|取得電子郵件訊息中的會議建議。|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;可為 null&gt;|取得電子郵件訊息或約會中的電話號碼。|
|`taskSuggestions`| Array.&lt;[TaskSuggestion](simple-types.md#tasksuggestion)&gt;| &lt;可為 null&gt;|取得電子郵件訊息或約會中的工作建議。|
|`urls`| Array.&lt;String&gt;| &lt;可為 null&gt;|取得電子郵件訊息或約會中的網際網路 URL。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
#### LocalClientTime

表示本機用戶端時區的日期和時間。僅限閱讀模式。

##### 屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`month`| 數字|代表月份的整數值，從 0 (代表 1 月) 開始到 11 (代表 12 月)。|
|`date`| 數字|代表月份日期的整數值。|
|`year`| 數字|代表年份的整數值。|
|`hours`| 數字|代表 24 小時制之小時的整數值。|
|`minutes`| 數字|代表分鐘的整數值。|
|`seconds`| 數字|代表秒的整數值。|
|`milliseconds`| 數字|代表毫秒的整數值。|
|`timezoneOffset`| 數字|代表本地時區與 UTC 之分鐘數差異的整數值。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
#### MeetingSuggestion

代表項目中的建議會議。僅限閱讀模式。

針對使用中項目呼叫 [`meetingSuggestions`](simple-types.md#entities) 或 [`Entities`](Office.context.mailbox.item.md#getentities--entities) 方法時，系統會在傳回之 [`getEntities`](Office.context.mailbox.item.md#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) 物件的 `getEntitiesByType` 屬性中傳回電子郵件訊息內的建議會議清單。

`start` 和 `end` 值是 Date 物件的字串表示，其包含建議之會議的開始和結束日期和時間。這些值採用針對目前使用者指定的預設時區。

##### 屬性：

|名稱| 類型	| 說明|
|---|---|---|
|`attendees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|取得建議會議的出席者。|
|`end`| String|取得建議會議的結束日期和時間。|
|`location`| String|取得建議會議的位置。|
|`meetingString`| String|取得已識別為會議建議的字串。|
|`start`| String|取得建議會議的開始日期和時間。|
|`subject`| String|取得建議會議的主旨。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
#### PhoneNumber

代表項目中識別的電話號碼。僅限閱讀模式。

含有電子郵件中電話號碼的 `PhoneNumber` 物件陣列，當您針對選定項目呼叫 [`phoneNumbers`](simple-types.md#entities) 方法時，系統會在傳回之 [`Entities`](Office.context.mailbox.item.md#getentities--entities) 物件的 `getEntities` 屬性內傳回該物件陣列。

##### 類型：

*   物件

##### 屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`originalPhoneString`| 字串|取得項目中識別為電話號碼的文字。|
|`phoneString`| 字串|取得包含電話號碼的字串。這個字串只含有電話號碼的數字，不含括號和連字號等字元 (如果原始項目有這些字元的話)。|
|`type`| String|取得識別電話號碼類型的字串：`Home`、`Work`、`Mobile`、`Unspecified`。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
#### TaskSuggestion

代表項目中識別的建議工作。僅限閱讀模式。

針對使用中項目呼叫 [`taskSuggestions`](simple-types.md#entities) 或 [`Entities`](Office.context.mailbox.item.md#getentities--entities) 方法時，系統會在傳回之 [`Entities`][`getEntities`](Office.context.mailbox.item.md#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) 物件的 `getEntitiesByType` 屬性中傳回電子郵件訊息內的建議工作清單。

##### 屬性：

|名稱| 類型	| 說明|
|---|---|---|
|`assignees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|取得應接受建議工作指派的使用者。|
|`taskString`| String|取得識別為工作建議的項目文字。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|
