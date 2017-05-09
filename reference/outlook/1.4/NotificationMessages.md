

# <a name="notificationmessages"></a>NotificationMessages

## <a name="notificationmessages"></a>NotificationMessages

`NotificationMessages` 物件會作為項目的 [`notificationMessages`](Office.context.mailbox.item.md#notificationmessages-notificationmessages) 屬性傳回。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="methods"></a>方法

####  <a name="addasynckey-jsonmessage-options-callback"></a>addAsync(key, JSONmessage, [options], [callback])

在項目中新增通知。

每個訊息最多有 5 則通知。設定過多會傳回 `NumberOfNotificationMessagesExceeded` 錯誤。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`key`| 字串||開發人員指定的索引鍵，可用來參考這則通知訊息。日後開發人員可以使用它來修改此訊息。其長度不可超過 32 個字元。|
|`JSONmessage`| 物件||JSON 物件，其中包含要新增至項目中的通知訊息。它包含下列屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>描述</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>指定訊息類型。如果類型是 <code>ProgressIndicator</code> 或 <code>ErrorMessage</code>，系統會自動提供圖示且訊息不會持續。因此，圖示和持續性的屬性不適用於這些類型的訊息。包含它們將會導致 <code>ArgumentException</code>。如果類型是 <code>ProgressIndicator</code>，當動作完成時，開發人員應該移除或取代進度指示器。</td></tr><tr><td><code>icon</code></td><td>String</td><td>在 <code>Resource</code> 區段之資訊清單中定義的圖示參考。它會出現在資訊列區域中。它只適用於當類型是 <code>InformationalMessage</code> 時。針對不支援的類型指定這個參數會導致例外狀況。</td></tr><tr><td><code>message</code></td><td>String</td><td>通知訊息的文字。最大長度為 150 個字元。如果開發人員傳遞較長的字串，系統會擲出 <code>ArgumentOutOfRange</code> 例外狀況。</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>僅適用於當類型是 <code>InformationalMessage</code> 時。如果是 <code>true</code>，系統會保留訊息，直到此增益集移除訊息或使用者關閉為止。如果是 <code>false</code>，當使用者巡覽至不同的項目時，便會遭到移除。若是錯誤通知，系統會保留訊息，直到使用者看到一次為止。針對不支援的類型指定這個參數會擲出例外狀況。</td></tr></tbody></table>|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 |

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
// Create three notifications, each with a different key
Office.context.mailbox.item.notificationMessages.addAsync("progress", {
  type: "progressIndicator",
  message : "An add-in is processing this message."
});
Office.context.mailbox.item.notificationMessages.addAsync("information", {
  type: "informationalMessage",
  message : "The add-in processed this message.",
  icon : "iconid",
  persistent: false
});
Office.context.mailbox.item.notificationMessages.addAsync("error", {
  type: "errorMessage",
  message : "The add-in failed to process this message."
});
```

####  <a name="getallasyncoptions-callback"></a>getAllAsync([options], [callback])

傳回項目的所有索引鍵和訊息。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

成功完成時，`asyncResult.value` 屬性會包含 [`NotificationMessageDetails`](simple-types.md#notificationmessagedetails) 物件的陣列。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
// Get all notifications
Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
  if (asyncResult.status != "failed") {
    Office.context.mailbox.item.notificationMessages.replaceAsync( "notifications", {
      type: "informationalMessage",
      message : "Found " + asyncResult.value.length + " notifications.",
      icon : "iconid",
      persistent: false
    });
  }
});
```

####  <a name="removeasynckey-options-callback"></a>removeAsync(key, [options], [callback])

移除項目的通知訊息。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`key`| 字串||要移除之通知訊息的索引鍵。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

如果找不到索引鍵，系統會在 `KeyNotFound` 屬性中傳回 `asyncResult.error` 錯誤。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
// Remove a notification
Office.context.mailbox.item.notificationMessages.removeAsync("progress");
```

####  <a name="replaceasynckey-jsonmessage-options-callback"></a>replaceAsync(key, JSONmessage, [options], [callback])

將具有指定索引鍵的通知訊息取代為另一則訊息。

如果具有指定索引鍵的通知訊息不存在，`replaceAsync` 會新增通知。

##### <a name="parameters"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`key`| String||要取代之通知訊息的索引鍵。其長度不可超過 32 個字元。|
|`JSONmessage`| 物件||JSON 物件，其包含要取代現有通知訊息的新訊息。它包含下列屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>描述</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>指定訊息類型。如果類型是 <code>ProgressIndicator</code> 或 <code>ErrorMessage</code>，系統會自動提供圖示且訊息不會持續。因此，圖示和持續性的屬性不適用於這些類型的訊息。包含它們將會導致 <code>ArgumentException</code>。如果類型是 <code>ProgressIndicator</code>，當動作完成時，開發人員應該移除或取代進度指示器。</td></tr><tr><td><code>icon</code></td><td>String</td><td>在 <code>Resource</code> 區段之資訊清單中定義的圖示參考。它會出現在資訊列區域中。它只適用於當類型是 <code>InformationalMessage</code> 時。針對不支援的類型指定這個參數會導致例外狀況。</td></tr><tr><td><code>message</code></td><td>String</td><td>通知訊息的文字。最大長度為 150 個字元。如果開發人員傳遞較長的字串，系統會擲出 <code>ArgumentOutOfRange</code> 例外狀況。</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>僅適用於當類型是 <code>InformationalMessage</code> 時。如果是 <code>true</code>，系統會保留訊息，直到此增益集移除訊息或使用者關閉為止。如果是 <code>false</code>，當使用者巡覽至不同的項目時，便會遭到移除。若是錯誤通知，系統會保留訊息，直到使用者看到一次為止。針對不支援的類型指定這個參數會擲出例外狀況。</td></tr></tbody></table>|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 |

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
// Replace a notification with an informational notification
Office.context.mailbox.item.notificationMessages.replaceAsync("progress", {
  type: "informationalMessage",
  message : "The message was processed successfully.",
  icon : "iconid",
  persistent: false
});
```
