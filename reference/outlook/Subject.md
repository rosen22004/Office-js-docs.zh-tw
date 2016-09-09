

# 主旨

提供在 Outlook 增益集中取得及設定約會或訊息主旨的方法。

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|

### 方法

####  getAsync([options], callback)

取得約會或訊息的主旨。

`getAsync` 方法會向 Exchange 伺服器啟動非同步呼叫，以取得約會或訊息的主旨。

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

項目主旨是以 `asyncResult.value` 屬性中的字串等形式提供。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|
####  setAsync(subject, [options], [callback])

設定約會或訊息的主旨。

`setAsync` 方法會向 Exchange 伺服器啟動非同步呼叫，以設定約會或訊息的主旨。設定主旨會覆寫目前的主旨，不過會保留任何現有的前置詞 (如 "Fwd:" 或 "Re:")。

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`subject`| String||約會或訊息的主旨。字串限制在 255 個字元內。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>如果主旨設定失敗，`asyncResult.error` 屬性將包含錯誤碼。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>說明</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>subject</code> 參數的長度超過 255 個字元。</td></tr></tbody></table>|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|
