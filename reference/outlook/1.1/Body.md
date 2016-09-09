

# 本文

`body` 物件提供新增和更新訊息或約會內容的方法。它會在所選取項目的 `body` 屬性中傳回。

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

### 方法

####  getTypeAsync([options], [callback])

取得的值會指出內容是 HTML 或文字格式的值。

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

系統會將傳回的內容類型當做 `asyncResult.value` 屬性中其中一個 [CoercionType](Office.md#coerciontype-string) 值。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|
####  prependAsync(data, [options], [callback])

將指定內容新增至項目本文的開頭。

`prependAsync` 方法會將指定字串插入項目本文的開頭。呼叫 `prependAsync` 方法等同於呼叫 [`setSelectedDataAsync`](Body.md#setselecteddataasyncdata-options-callback) 方法，並將插入點置於本文內容的開頭。

在 HTML 標記中包含連結時，您可以將錨點 (`<a>`) 上的 `id` 屬性設定為 `LPNoLP`，藉此停用線上連結預覽。 例如：

```
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`data`| String||要插入本文開頭的字串。字串限制在 1,000,000 個字元內。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>說明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;選擇性&gt;</td><td>所需的本文格式。<code>data</code> 參數中的字串將會轉換為此格式。</td></tr><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>系統會在 `asyncResult.error` 屬性中提供任何發生的錯誤。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>說明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> 參數的長度超過 1,000,000 個字元。</td></tr></tbody></table>|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|
####  setSelectedDataAsync(data, [options], [callback])

將本文中的選取範圍內容取代為指定文字。

`setSelectedDataAsync` 方法會在項目本文中的游標位置插入指定字串；如果您已在編輯器中選取文字，它會取代選取的文字。如果游標從未置於項目本文中，或項目本文在 UI 中遺失焦點，系統會將字串插入本文內容的頂端。

在 HTML 標記中包含連結時，您可以將錨點 (`<a>`) 上的 `id` 屬性設定為 `LPNoLP`，藉此停用線上連結預覽。 例如：

```
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`data`| String||要插入本文的字串。字串限制在 1,000,000 個字元內。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>說明</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;選擇性&gt;</td><td>所需的本文格式。<code>data</code> 參數中的字串將會轉換為此格式。</td></tr><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>系統會在 `asyncResult.error` 屬性中提供任何發生的錯誤。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>說明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>data</code> 參數長度超過 1,000,000 個字元。</td></tr><tr><td><code>InvalidFormatError</code></td><td>本文類型已設定為 HTML 且 data 參數含有純文字。</td></tr></tbody></table>|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|適用的 Outlook 模式| 撰寫|
