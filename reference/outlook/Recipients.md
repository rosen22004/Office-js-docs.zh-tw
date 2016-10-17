

# <a name="recipients"></a>收件者

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|

### <a name="methods"></a>方法

####  <a name="addasync(recipients,-[options],-[callback])"></a>addAsync(recipients, [options], [callback])

新增收件者清單至現有約會或郵件的收件者。

`recipients` 參數可以是下列其中一個陣列：

*   字串包含 SMTP 電子郵件地址
*   `EmailUser` 物件
*   `EmailAddressDetails` 物件

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`recipients`| 陣列.&lt;(字串&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||會將收件者新增到收件者清單。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>如果收件者新增失敗，`asyncResult.error` 屬性將包含錯誤碼。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>描述</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>收件者數目超過 100 個項目。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="example"></a>範例

下列範例會建立 `EmailUser` 物件的陣列，並將它們新增至郵件的收件者。

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.to.addAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients added");
  }
});
```

####  <a name="getasync([options],-callback)"></a>getAsync([options], callback)

取得約會或郵件的收件者清單。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

當呼叫完成時，`asyncResult.value` 屬性會包含 [`EmailAddressDetails`](simple-types.md#emailaddressdetails) 物件的陣列。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="example"></a>範例

下列範例會取得會議的列席者。

```js
Office.context.mailbox.item.optionalAttendees.getAsync(function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    var msg = "";
    result.value.forEach(function(recip, index) {
      msg = msg + recip.displayName + " (" + recip.emailAddress + ");";
    });
    showMessage(msg);
  }
});
```

####  <a name="setasync(recipients,-[options],-callback)"></a>setAsync(recipients, [options], callback)

設定約會或郵件的收件者清單。

`setAsync` 方法會覆寫目前的收件者清單。

`recipients` 參數可以是下列其中一個陣列：

*   字串包含 SMTP 電子郵件地址
*   `EmailUser` 物件
*   `EmailAddressDetails` 物件

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`recipients`| 陣列.&lt;(字串&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||會將收件者新增到收件者清單。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>如果收件者設定失敗，`asyncResult.error` 屬性將含有代碼，其代表在新增資料時發生的任何錯誤。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>描述</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>收件者數目超過 100 個項目。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="example"></a>範例

下列範例會建立 `EmailUser` 物件的陣列，並以陣列取代郵件的副本收件者。

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.cc.setAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients overwritten");
  }
});
```
