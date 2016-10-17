

# <a name="time"></a>時間

`Time` 物件會作為撰寫模式中約會的 [`start`](Office.context.mailbox.item.md#start-datetime) 或 [`end`](Office.context.mailbox.item.md#end-datetime) 屬性傳回。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|

### <a name="methods"></a>方法

####  <a name="getasync([options],-callback)"></a>getAsync([options], callback)

取得約會的開始或結束時間。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

日期和時間依 `asyncResult.value` 屬性中的 Date 物件提供。值為國際標準時間 (UTC)。您可以使用 [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法將 UTC 時間轉換成本機用戶端時間。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|
####  <a name="setasync(datetime,-[options],-[callback])"></a>setAsync(dateTime, [options], [callback])

設定約會的開始或結束時間。

如果針對 [`setAsync`](Office.context.mailbox.item.md#start-datetime) 屬性呼叫 `start` 方法，系統會調整 [`end`](Office.context.mailbox.item.md#end-datetime) 屬性以維持先前設定之約會的持續時間。如果針對 `setAsync` 屬性呼叫 `end` 方法，系統會延長約會的持續時間以符合新的結束時間。

時間必須是 UTC；您可以使用 [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) 方法取得正確的 UTC 時間。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`dateTime`| 日期||採用國際標準時間 (UTC) 的 Date 物件。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>如果日期和時間設定失敗，`asyncResult.error` 屬性將含有錯誤碼。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>描述</th></tr></thead><tbody><tr><td><code>InvalidEndTime</code></td><td>約會結束時間比約會開始時間早。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|

##### <a name="example"></a>範例

下列範例會設定約會的開始時間。

```js
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
