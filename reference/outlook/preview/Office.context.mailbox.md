

# <a name="mailbox"></a>信箱

## [Office](Office.md)[.context](Office.context.md). mailbox

提供用於存取 Microsoft Outlook 和網頁型 Outlook 的 Outlook 增益集物件模型。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="namespaces"></a>命名空間

[診斷](Office.context.mailbox.diagnostics.md):提供診斷資訊給 Outlook 增益集。

[項目](Office.context.mailbox.item.md):提供用來在 Outlook 增益集中存取郵件或約會的方法和屬性。

[userProfile](Office.context.mailbox.userProfile.md):提供在 Outlook 增益集中使用者的相關資訊。</dd>

### <a name="members"></a>成員

#### <a name="ewsurl-string"></a>ewsUrl︰字串

取得這個電子郵件帳戶的 Exchange Web 服務 (EWS) 端點的 URL。僅限閱讀模式。

> **附註：**iOS 版 Outlook 或 Android 版 Outlook 不支援這個成員。

遠端服務可以使用 `ewsUrl` 的值，來將 EWS 呼叫至使用者信箱。例如，您可以建立遠端服務以[從選取項目取得附件](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx)。

您的應用程式必須在其資訊清單中具有指定的 **ReadItem** 權限，才能在讀取模式中呼叫 `ewsUrl` 成員。

在撰寫模式中，您必須先呼叫 [`saveAsync`](Office.context.mailbox.item#saveAsync) 方法，才可以使用 `ewsUrl` 成員。您的應用程式必須具有 **ReadWriteItem** 權限以呼叫 `saveAsync` 方法。

##### <a name="type"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

#### <a name="resturl-string"></a>restUrl :String

取得此電子郵件帳戶的 REST 端點的 URL。

可使用 `restUrl` 的值，來對使用者信箱產生 [REST API](https://dev.outlook.com/restapi/reference) 呼叫。

您的應用程式必須在其資訊清單中具有指定的 **ReadItem** 權限，才能在讀取模式中呼叫 `restUrl` 成員。

在撰寫模式中，您必須先呼叫 [`saveAsync`](Office.context.mailbox.item#saveAsync) 方法，才可以使用 `restUrl` 成員。您的應用程式必須具有 **ReadWriteItem** 權限以呼叫 `saveAsync` 方法。

##### <a name="type"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.5 |
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="methods"></a>方法

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

新增支援事件的事件處理常式。

目前唯一支援的事件類型是 `Office.EventType.ItemChanged`，當使用者選取新項目時會叫用此類型。這個事件由實作可釘選工作窗格的增益集使用，且此事件讓增益集可根據目前選取的項目重新整理工作窗格 UI。

##### <a name="parameters"></a>參數：

| 名稱 | 類型	 | 屬性 | 描述 |
|---|---|---|---|
| `eventType` | [Office.EventType](Office.md#EventType) || 應叫用處理常式的事件。 |
| `handler` | 函數 || 若要處理事件的函數。函數必須接受單一參數，也就是物件常值。參數上的 `type` 屬性將會符合傳遞至 `addHandlerAsync` 的 `eventType` 參數。 |
| `options` | 物件 | &lt;選擇性&gt; | 物件常值包含下列一或多個屬性。 |
| `options.asyncContext` | 物件 | &lt;選擇性&gt; | 開發人員可提供任何他們想要在回呼方法中存取的物件。 |
| `callback` | 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.5 |
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem |
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

將 REST 的項目 ID 轉換為 EWS 格式。

> **附註：**iOS 版 Outlook 或 Android 版 Outlook 不支援這個方法。

透過 REST API 擷取的項目 ID (例如 [Outlook 郵件 API](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) 或 [Microsoft Graph](http://graph.microsoft.io/)) 使用的格式與 Exchange Web 服務 (EWS) 所使用的格式不同。`convertToEwsId` 方法將 REST 格式的 ID 轉換成 EWS 的適當格式。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`itemId`| 字串|針對 Outlook REST API 格式化的項目 ID|
|`restVersion`| [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#restversion)|值，指出用來擷取項目 ID 的 Outlook REST API 版本。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="returns"></a>傳回：

類型：字串

##### <a name="example"></a>範例

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimesimple-typesmdlocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](simple-types.md#localclienttime)}

取得包含在本機用戶端時間中時間資訊的字典。

用於 Outlook 或 Outlook Web App 中郵件應用程式中的日期和時間，可以使用不同時區。Outlook 會使用用戶端電腦的時區；Outlook Web App 則會使用 Exchange 系統管理中心 (EAC) 上所設定的時區。您應該處理日期和時間值，如此使用者介面上所顯示的值會永遠和使用者所預期的時區一致。

如果郵件應用程式是在 Outlook 中執行，`convertToLocalClientTime` 方法會傳回字典物件，內含設定為用戶端電腦時區的值。如果郵件應用程式是在 Outlook Web App 中執行，`convertToLocalClientTime` 方法會傳回字典物件，內含設定為 EAC中指定時區的值。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`timeValue`| 日期|日期物件|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="returns"></a>傳回：

類型：[LocalClientTime](simple-types.md#localclienttime)

####  <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

將 EWS 的項目 ID 轉換為 REST 格式。

> **附註：**iOS 版 Outlook 或 Android 版 Outlook 不支援這個方法。

透過 EWS 或 `itemId` 屬性擷取的項目 ID，使用的格式與 REST API 所使用的格式不同 (例如 [Outlook 郵件 API](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) 或 [Microsoft Graph](http://graph.microsoft.io/))。`convertToRestId` 方法將 EWS 格式的 ID 轉換成 REST 的適當格式。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`itemId`| 字串|針對 Exchange Web 服務 (EWS) 格式化的項目 ID|
|`restVersion`| [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#restversion)|值，指出用於轉換 ID 的 Outlook REST API 版本。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="returns"></a>傳回：

類型：字串

##### <a name="example"></a>範例

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

取得來自字典包含時間資訊的日期物件。

`convertToUtcClientTime` 方法會將包含本機日期和時間的字典，轉換至具有本機正確日期和時間值的日期物件。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`input`| [LocalClientTime](simple-types.md#localclienttime)|要轉換的本機時間值。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="returns"></a>傳回：

以 UTC 表示時間的日期物件。

<dl class="param-type">

<dt>
類型</dt>


<dd>日期</dd>

</dl>

####  <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

顯示現有的行事曆約會。

> **附註：**iOS 版 Outlook 或 Android 版 Outlook 不支援這個方法。

`displayAppointmentForm` 方法會在桌面上的新視窗或是在行動裝置上的對話方塊，開啟現有的行事曆約會。

在 Mac 版 Outlook 中，您可以使用這個方法，來顯示不屬於週期性系列的單一約會，或是週期性系列的主約會；但您無法顯示系列的執行個體。這是因為在 Mac 版 Outlook 中，您無法存取週期性系列的執行個體屬性 (包括項目識別碼)。

在 Outlook Web App 中，唯有當表單的本文是小於或等於 32 KB 的字元數，這個方法才會開啟指定的表單。

如果指定的項目識別碼無法識別現有的約會，會在用戶端電腦或裝置上開啟空白窗格，而且不會傳回錯誤訊息。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`itemId`| 字串|現有行事曆約會的 Exchange Web 服務 (EWS) 識別碼。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

顯示現有的郵件。

> **附註：**iOS 版 Outlook 或 Android 版 Outlook 不支援這個方法。

`displayMessageForm` 方法會在桌面上的新視窗，或是在行動裝置上的對話方塊，開啟現有郵件。

在 Outlook Web App 中，唯有當表單的本文是小於或等於 32 KB 的字元數，這個方法才會開啟指定的表單。

如果指定的項目識別碼無法識別現有的郵件，便不會在用戶端電腦或裝置上顯示郵件，而且不會傳回錯誤訊息。

請勿將 `displayMessageForm` 與表示約會的 `itemId` 一起使用。使用 `displayAppointmentForm` 方法來顯示現有約會，並使用 `displayNewAppointmentForm` 來顯示表單，以建立新的約會。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`itemId`| 字串|現有郵件的 Exchange Web 服務 (EWS) 識別碼。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

顯示表單來建立新的行事曆約會。

> **附註：**iOS 版 Outlook 或 Android 版 Outlook 不支援這個方法。

`displayNewAppointmentForm` 方法會開啟表單，讓使用者建立新的約會或會議。如果指定參數，會將參數的內容自動填入約會表單欄位。

在 Outlook Web App 和 OWA for Devices 中，這個方法永遠會顯示帶有出席者欄位的表單。如果您未指定任何出席者作為輸入引數，則方法會顯示帶有 [儲存]**** 按鈕的表單。如果您已指定出席者，表單會包括出席者和 [傳送]**** 按鈕。

在 Outlook 豐富型用戶端和 Outlook RT 中，如果您在 `requiredAttendees`、`optionalAttendees` 或 `resources` 參數中指定任何出席者或資源，這個方法會顯示帶有 [傳送]**** 按鈕的會議表單。如果您不指定任何收件者，這個方法會顯示帶有 [儲存 & 關閉]**** 按鈕的約會表單。

如果任何參數超過指定的大小限制，或指定未知的參數名稱，則會拋出例外狀況。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 描述|
|---|---|---|
|`parameters`| 物件|描述新的約會的參數字典。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>描述</th></tr></thead><tbody><tr><td><code>requiredAttendees</code></td><td>陣列。&lt;字串&gt; &#124; 陣列。&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>包含電子郵件地址的字串陣列，或包含約會的每個必要出席者的 <code>EmailAddressDetails</code> 物件的陣列。這個陣列限制最多為 100 個項目。</td></tr><tr><td><code>optionalAttendees</code></td><td>陣列。&lt;字串&gt; &#124; 陣列。&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>包含電子郵件地址的字串陣列，或包含約會的每個選擇性出席者的 EmailAddressDetails 物件的陣列。這個陣列限制最多為 100 個項目。</td></tr><tr><td><code>start</code></td><td>日期</td><td>日期物件指定約會的開始日期和時間。</td></tr><tr><td><code>end</code></td><td>日期</td><td>日期物件指定約會的結束日期和時間。</td></tr><tr><td><code>location</code></td><td>字串</td><td>字串，包含約會地點。字串限制在最多 255 個字元。</td></tr><tr><td><code>resources</code></td><td>Array.&lt;String&gt;</td><td>包含約會所需資源的字串陣列。這個陣列限制最多為 100 個項目。</td></tr><tr><td><code>subject</code></td><td>字串</td><td>字串，包含約會主旨。字串限制在最多 255 個字元。</td></tr><tr><td><code>body</code></td><td>字串</td><td>約會訊息的本文。本文內容檔案大小的上限為 32 KB。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 讀取|

##### <a name="example"></a>範例

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

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync([options], callback)

取得字串，其中包含用來呼叫 REST APIs 或 Exchange Web 服務的權杖。

`getCallbackTokenAsync` 方法會以非同步方式呼叫，以便從裝載使用者信箱的 Exchange Server 取得不透明權杖。回呼權杖的存留期為 5 分鐘。

> **附註：**如果可能的話，建議增益集使用 REST API，而不是 Exchange Web 服務。 

**REST 權杖**

要求 REST 權杖 (`options.isRest = true`) 時，產生的權杖無法運作來驗證 Exchange Web 服務呼叫。權杖的範圍將會受限於僅可唯讀存取目前項目和其附件，除非增益集已在其資訊清單中指定 [`ReadWriteMailbox` ](../../docs/add-ins/outlook/understanding-outlook-add-in-permissions#readwritemailbox-permission) 權限。如果 `ReadWriteMailbox` 權限已指定，則產生的權杖將會授與對郵件、行事曆及連絡人的讀取/寫入存取權，其中包括傳送郵件的能力。

增益集應使用 `restUrl` 屬性來決定在進行 API 呼叫時所使用正確的 URL。

**EWS 權杖**

要求 REST 權杖 (`options.isRest = false`) 時，產生的權杖無法運作來驗證 REST API 呼叫。權杖的範圍將會受限於存取目前的項目。

增益集應使用 `ewsUrl` 屬性來決定進行 EWS 呼叫時正確的 URL。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
| `options` | 物件 | &lt;選擇性&gt; | 物件常值包含下列一或多個屬性。 |
| `options.isRest` | 布林值 |  &lt;選用&gt; | 決定提供的記號是否會用於 Outlook REST API 或 Exchange Web 服務。預設值是 `false`。 |
| `options.asyncContext` | 物件 |  &lt;選用&gt; | 傳遞至非同步方法的任何狀態資料。 |
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。權限是在 `asyncResult.value` 屬性中提供作為字串。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.5 |
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫和讀取|

##### <a name="example"></a>範例

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

取得包含用來從 Exchange Server 取得附件或項目權杖的字串。

`getCallbackTokenAsync` 方法會以非同步方式呼叫，以便從裝載使用者信箱的 Exchange Server 取得不透明權杖。回呼權杖的存留期為 5 分鐘。

您可以將權杖和附件識別碼或項目識別碼傳遞至協力廠商系統。協力廠商系統使用權杖作為承載者授權權杖，以呼叫此 Exchange Web 服務 (EWS) [GetAttachment](https://msdn.microsoft.com/en-us/library/office/aa494316.aspx) 或 [GetItem](https://msdn.microsoft.com/en-us/library/office/aa565934.aspx) 作業來傳回附件或項目。例如，您可以建立遠端服務以[從選取項目取得附件](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx)。

您的應用程式必須在其資訊清單中具有指定的 **ReadItem** 權限，才能在讀取模式中呼叫 `getCallbackTokenAsync` 方法。

在撰寫模式中，您必須呼叫 [`saveAsync`](Office.context.mailbox.item#saveAsync) 方法來取得要傳遞至 `getCallbackTokenAsync` 方法的項目識別碼。您的應用程式必須具有 **ReadWriteItem** 權限以呼叫 `saveAsync` 方法。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。權限是在 `asyncResult.value` 屬性中提供作為字串。|
|`userContext`| 物件| &lt;選用&gt;|傳遞至非同步方法的任何狀態資料。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫和讀取|

##### <a name="example"></a>範例

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

取得用來識別使用者及 Office 增益集的權杖。

`getUserIdentityTokenAsync` 方法會傳回權杖，可用來識別和[驗證增益集和協力廠商系統的使用者](https://msdn.microsoft.com/EN-US/library/office/fp179828.aspx)。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

權杖是在 `asyncResult.value` 屬性中提供為字串。| |`userContext`| 物件 | &lt;選擇性&gt;|傳遞至非同步方法的任何狀態資料。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

在裝載使用者信箱的 Exchange Server 上，對 Exchange Web 服務 (EWS) 提出非同步要求。

> **附註：**iOS 版 Outlook 或 Android 版 Outlook 不支援這個方法。

`makeEwsRequestAsync` 方法會代表增益集將 EWS 要求傳送到 Exchange。

您無法以 `makeEwsRequestAsync` 方法要求資料夾關聯項目。

XML 要求必須指定 UTF-8 編碼。

```
<?xml version="1.0" encoding="utf-8"?>
```

您的增益集必須具有 **ReadWriteMailbox** 權限以使用 `makeEwsRequestAsync` 方法。有關如何使用 **ReadWriteMailbox** 權限和以 `makeEwsRequestAsync` 方法呼叫 EWS 作業的詳細資訊，請參閱[指定郵件增益集存取至使用者信箱的權限](../../../docs/outlook/understanding-outlook-add-in-permissions.md)。

**附註**：伺服器管理員必須在用戶端存取伺服器 EWS 目錄上，將 `OAuthAuthentication` 設定為 true，以便啟用 `makeEwsRequestAsync` 方法來提出 EWS 要求。

#### <a name="version-differences"></a>版本差異

當在郵件應用程式內 (於早於 15.0.4535.1004 的 Outlook 版本中執行) 使用 `makeEwsRequestAsync` 方法，您應該將編碼的值設定為 `ISO-8859-1`。

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

當郵件應用程式在 Outlook 網頁版中執行時，您不需要設定編碼值。您可以使用 mailbox.diagnostics.hostName 屬性，來決定郵件應用程式是否正在 Outlook 或在 Outlook 網頁版中執行。您可以使用 mailbox.diagnostics.hostVersion 屬性，來決定所執行的 Outlook 版本。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`data`| 字串||EWS 要求。|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

將 EWS 呼叫的 XML 結果在 `asyncResult.value` 屬性中提供作為字串。如果結果的大小超過 1 MB，反而會傳回錯誤訊息。| |`userContext`| 物件 | &lt;選擇性&gt;|傳遞至非同步方法的任何狀態資料。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteMailbox|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

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