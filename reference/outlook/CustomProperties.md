

# CustomProperties

`CustomProperties` 物件代表特定項目所特有的和適用於 Outlook 的郵件增益集所特有的自訂屬性。 例如，對於啟動郵件增益集的現行電子郵件訊息，您可能會需要增益集儲存電子郵件訊息特有的資料。 如果使用者在未來重新造訪相同的訊息並再次啟動郵件增益集，增益集將能擷取已儲存為自訂屬性的資料。

由於 Mac 版 Outlook 不會快取自訂屬性，所以如果使用者的網路連線中斷，郵件增益集便無法存取其自訂屬性。

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

### 範例

下列範例示範如何使用 `loadCustomPropertiesAsync` 方法以非同步方式載入目前項目特有的自訂屬性。 範例中也示範如何使用 [`saveAsync`](#saveasync) 方法，將這些屬性儲存回伺服器。 載入自訂屬性之後，程式碼範例會使用 [`get`](CustomProperties.md#getname--string) 方法讀取自訂屬性 `myProp`，使用 [`set`](CustomProperties.md#setname-value) 方法撰寫自訂屬性 `otherProp`，然後最後呼叫 `saveAsync` 方法，來儲存自訂屬性。

```
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var mailbox = Office.context.mailbox;
    mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
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

### 方法

####  get(name) → {String}

傳回指定之自訂屬性的值。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`name`| 字串|要傳回之自訂屬性的名稱。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### 傳回：

指定之自訂屬性的值。

<dl class="param-type">

<dt>
類型</dt>


<dd>字串</dd>

</dl>

####  remove(name)

從自訂屬性集合中移除指定屬性。

若要永久移除屬性，您必須呼叫 `CustomProperties` 物件的 [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) 方法。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`name`| 字串|要移除的屬性名稱。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
####  saveAsync([callback], [asyncContext])

將項目特有的自訂屬性儲存至伺服器。

您必須呼叫 `saveAsync` 方法，才能保存以 [`set`](CustomProperties.md#setname-value) 方法或 `CustomProperties` 物件之 [`remove`](CustomProperties.md#removename) 方法所做的變更。 儲存動作採用非同步方式。

讓回呼函數檢查及處理來自 `saveAsync` 的錯誤是實用的做法。 特別是，當使用者在讀取表單中處於連線狀態時會啟動讀取增益集，不過之後使用者的連線可能會中斷。 如果增益集在中斷連線狀態下呼叫 `saveAsync`，`saveAsync` 便會傳回錯誤。 您的回呼方法應該要能相應地處理這個錯誤。

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 |
|`asyncContext`| 物件| &lt;選用&gt;|傳遞至回呼方法的任何狀態資料。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### 範例

以下 JavaScript 程式碼範例示範如何以非同步方式使用 `loadCustomPropertiesAsync` 方法，載入目前項目特有的自訂屬性，以及使用 [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) 方法將這些屬性儲存回伺服器。 載入自訂屬性之後，程式碼範例會使用 [`get`](#get) 方法讀取自訂屬性 `myProp`，使用 [`set`](CustomProperties.md#setname-value) 方法撰寫自訂屬性 `otherProp`，然後最後呼叫 `saveAsync` 方法，來儲存自訂屬性。

```
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
  if (asyncResult.status == Office.AsyncResultStatus.Failed){
    write(asyncResult.error.message);
  }
  else {
    // Async call to save custom properties completed.
    // Proceed to do the appropriate for your add-in.
  }
}

// Writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  set(name, value)

將指定屬性設定為指定值。

`set` 方法會將指定屬性設定為指定值。 您必須使用 [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) 方法將屬性儲存至伺服器。

如果指定的屬性不存在，`set` 方法會建立新屬性，否則會將現有的值取代為新的值。 `value` 參數可以是任何類型，不過系統一律會以字串形式將它傳遞至伺服器。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`name`| 字串|要設定的屬性名稱。|
|`value`| 物件|要設定的屬性值。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|
