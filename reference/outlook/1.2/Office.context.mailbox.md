

# 信箱

## [Office](Office.md)[.context](Office.context.md). 信箱

提供用於存取 Microsoft Outlook 和網頁型 Outlook 的 Outlook 增益集物件模型。

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

### 命名空間

[診斷](Office.context.mailbox.diagnostics.md):提供診斷資訊給 Outlook 增益集。

[項目](Office.context.mailbox.item.md):提供用來在 Outlook 增益集中存取郵件或約會的方法和屬性。

[userProfile](Office.context.mailbox.userProfile.md):提供在 Outlook 增益集中使用者的相關資訊。

### 成員

#### ewsUrl︰字串

取得這個電子郵件帳戶的 Exchange Web 服務 (EWS) 端點的 URL。 僅限閱讀模式。

遠端服務可以使用 `ewsUrl` 的值，來將 EWS 呼叫至使用者信箱。 例如，您可以建立遠端服務以[從選取項目取得附件](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx)。

##### 類型：

*   字串

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

### 方法

####  convertToLocalClientTime(timeValue) → {[LocalClientTime](simple-types.md#localclienttime)}

取得包含在本機用戶端時間中時間資訊的字典。

用於 Outlook 或 Outlook Web App 中郵件應用程式中的日期和時間，可以使用不同時區。Outlook 會使用用戶端電腦的時區；Outlook Web App 則會使用 Exchange 系統管理中心 (EAC) 上所設定的時區。您應該處理日期和時間值，如此使用者介面上所顯示的值會永遠和使用者所預期的時區一致。

如果郵件應用程式是在 Outlook 中執行，`convertToLocalClientTime` 方法會傳回字典物件，內含設定為用戶端電腦時區的值。 如果郵件應用程式是在 Outlook Web App 中執行，`convertToLocalClientTime` 方法會傳回字典物件，內含設定為 EAC中指定時區的值。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`timeValue`| 日期|日期物件|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### 傳回：

類型：[LocalClientTime](simple-types.md#localclienttime)

####  convertToUtcClientTime(input) → {Date}

取得來自字典包含時間資訊的日期物件。

`convertToUtcClientTime` 方法會將包含本機日期和時間的字典，轉換至具有本機正確日期和時間值的日期物件。

##### 參數：

|名稱| 類型	| 說明|
|---|---|---|
|`input`| [LocalClientTime](simple-types.md#localclienttime)|要轉換的本機時間值。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### 傳回：

以 UTC 表示時間的日期物件。

<dl class="param-type">

<dt>
類型</dt>


<dd>日期</dd>

</dl>

####  displayAppointmentForm(itemId)

顯示現有的行事曆約會。

`displayAppointmentForm` 方法會在桌面上的新視窗或是在行動裝置上的對話方塊，開啟現有的行事曆約會。

在 Mac 版 Outlook 中，您可以使用這個方法，來顯示不屬於週期性系列的單一約會，或是週期性系列的主約會；但您無法顯示系列的執行個體。這是因為在 Mac 版 Outlook 中，您無法存取週期性系列的執行個體屬性 (包括項目識別碼)。

在 Outlook Web App 中，唯有當表單的本文是小於或等於 32 KB 的字元數，這個方法才會開啟指定的表單。

如果指定的項目識別碼無法識別現有的約會，會在用戶端電腦或裝置上開啟空白窗格，而且不會傳回錯誤訊息。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`itemId`| 字串|現有行事曆約會的 Exchange Web 服務 (EWS) 識別碼。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### 範例

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  displayMessageForm(itemId)

顯示現有的郵件。

`displayMessageForm` 方法會在桌面上的新視窗，或是在行動裝置上的對話方塊，開啟現有郵件。

在 Outlook Web App 中，唯有當表單的本文是小於或等於 32 KB 的字元數，這個方法才會開啟指定的表單。

如果指定的項目識別碼無法識別現有的郵件，便不會在用戶端電腦或裝置上顯示郵件，而且不會傳回錯誤訊息。

請勿將 `displayMessageForm` 與表示約會的 `itemId` 一起使用。 使用 `displayAppointmentForm` 方法來顯示現有約會，並使用 `displayNewAppointmentForm` 來顯示表單，以建立新的約會。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`itemId`| 字串|現有郵件的 Exchange Web 服務 (EWS) 識別碼。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### 範例

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### displayNewAppointmentForm(parameters)

顯示表單來建立新的行事曆約會。

`displayNewAppointmentForm` 方法會開啟表單，讓使用者建立新的約會或會議。 如果指定參數，會將參數的內容自動填入約會表單欄位。

在 Outlook Web App 和 OWA for Devices 中，這個方法永遠會顯示帶有出席者欄位的表單。 如果您未指定任何出席者作為輸入引數，則方法會顯示帶有 [儲存]**** 按鈕的表單。 如果您已指定出席者，表單會包括出席者和 [傳送]**** 按鈕。

在 Outlook 豐富型用戶端和 Outlook RT 中，如果您在 `requiredAttendees`、`optionalAttendees` 或 `resources` 參數中指定任何出席者或資源，這個方法會顯示帶有 [傳送]**** 按鈕的會議表單。 如果您不指定任何收件者，這個方法會顯示帶有 [儲存 & 關閉]**** 按鈕的約會表單。

如果任何參數超過指定的大小限制，或指定未知的參數名稱，則會拋出例外狀況。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`parameters`| 物件|描述新的約會的參數字典。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>說明</th></tr></thead><tbody><tr><td><code>requiredAttendees</code></td><td>陣列。&lt;字串&gt; &#124; 陣列。&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>包含電子郵件地址的字串陣列，或包含約會的每個必要出席者的 <code>EmailAddressDetails</code> 物件的陣列。 這個陣列限制最多為 100 個項目。</td></tr><tr><td><code>optionalAttendees</code></td><td>陣列。&lt;字串&gt; &#124; 陣列。&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>包含電子郵件地址的字串陣列，或包含約會的每個選擇性出席者的 EmailAddressDetails 物件的陣列。 這個陣列限制最多為 100 個項目。</td></tr><tr><td><code>start</code></td><td>日期</td><td>日期物件指定約會的開始日期和時間。</td></tr><tr><td><code>end</code></td><td>日期</td><td>日期物件指定約會的結束日期和時間。</td></tr><tr><td><code>location</code></td><td>字串</td><td>字串，包含約會地點。 字串限制在最多 255 個字元。</td></tr><tr><td><code>resources</code></td><td>Array.&lt;String&gt;</td><td>包含約會所需資源的字串陣列。 這個陣列限制最多為 100 個項目。</td></tr><tr><td><code>subject</code></td><td>字串</td><td>字串，包含約會主旨。 字串限制在最多 255 個字元。</td></tr><tr><td><code>body</code></td><td>字串</td><td>約會訊息的本文。 本文內容檔案大小的上限為 32 KB。</td></tr></tbody></table>|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### 範例

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### getCallbackTokenAsync(callback, [userContext])

取得包含用來從 Exchange Server 取得附件或項目權杖的字串。

`getCallbackTokenAsync` 方法會以非同步方式呼叫，以便從裝載使用者信箱的 Exchange Server 取得不透明權杖。 回呼權杖的存留期為 5 分鐘。

您可以將權杖和附件識別碼或項目識別碼傳遞至協力廠商系統。 協力廠商系統使用權杖作為承載者授權權杖，以呼叫此 Exchange Web 服務 (EWS) [GetAttachment](https://msdn.microsoft.com/en-us/library/office/aa494316.aspx) 或 [GetItem](https://msdn.microsoft.com/en-us/library/office/aa565934.aspx) 作業來傳回附件或項目。 例如，您可以建立遠端服務以[從選取項目取得附件](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx)。

您的應用程式必須在資訊清單中具有所指定的 **ReadItem** 權限，來呼叫 `getCallbackTokenAsync` 方法。

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

權杖是在 `asyncResult.value` 屬性中提供為字串。| |`userContext`| 物件 | &lt;選擇性&gt;|傳遞至非同步方法的任何狀態資料。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### 範例

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  getUserIdentityTokenAsync(callback, [userContext])

取得用來識別使用者及 Office 增益集的權杖。


  `getUserIdentityTokenAsync` 方法會傳回權杖，可用來識別和[驗證增益集和協力廠商系統的使用者](https://msdn.microsoft.com/EN-US/library/office/fp179828.aspx)。

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

權杖是在 `asyncResult.value` 屬性中提供為字串。| |`userContext`| 物件 | &lt;選擇性&gt;|傳遞至非同步方法的任何狀態資料。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### 範例

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  makeEwsRequestAsync(data, callback, [userContext])

在裝載使用者信箱的 Exchange Server 上，對 Exchange Web 服務 (EWS) 提出非同步要求。

`makeEwsRequestAsync` 方法會代表增益集將 EWS 要求傳送到 Exchange。

您無法以 `makeEwsRequestAsync` 方法要求資料夾關聯項目。

XML 要求必須指定 UTF-8 編碼。

```
<?xml version="1.0" encoding="utf-8"?>
```

您的增益集必須具有 **ReadWriteMailbox** 權限以使用 `makeEwsRequestAsync` 方法。 有關如何使用 **ReadWriteMailbox** 權限和以 `makeEwsRequestAsync` 方法呼叫 EWS 作業的詳細資訊，請參閱[指定郵件增益集存取至使用者信箱的權限](../../../docs/outlook/understanding-outlook-add-in-permissions.md)。

**附註**：伺服器管理員必須在用戶端存取伺服器 EWS 目錄上，將 `OAuthAuthentication` 設定為 true，以便啟用 `makeEwsRequestAsync` 方法來提出 EWS 要求。

#### 版本差異

當在郵件應用程式內 (於早於 15.0.4535.1004 的 Outlook 版本中執行) 使用 `makeEwsRequestAsync` 方法，您應該將編碼的值設定為 `ISO-8859-1`。

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

當郵件應用程式在 Outlook 網頁版中執行時，您不需要設定編碼值。您可以使用 mailbox.diagnostics.hostName 屬性，來決定郵件應用程式是否正在 Outlook 或在 Outlook 網頁版中執行。您可以使用 mailbox.diagnostics.hostVersion 屬性，來決定所執行的 Outlook 版本。

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`data`| 字串||EWS 要求。|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

將 EWS 呼叫的 XML 結果在 `asyncResult.value` 屬性中提供作為字串。 如果結果的大小超過 1 MB，反而會傳回錯誤訊息。| |`userContext`| 物件 | &lt;選擇性&gt;|傳遞至非同步方法的任何狀態資料。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteMailbox|
|適用的 Outlook 模式| 撰寫或讀取|

##### 範例

下列範例會呼叫 `makeEwsRequestAsync`，以便使用 `GetItem` 作業來取得項目主旨。

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```
