

# <a name="location"></a>地點

提供在 Outlook 增益集中取得及設定會議地點的方法。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|

### <a name="methods"></a>方法

####  <a name="getasync([options],-callback)"></a>getAsync([options], callback)

取得約會的地點。

`getAsync` 方法會向 Exchange 伺服器啟動非同步呼叫，以取得約會的地點。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數||當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。

約會的地點以 `asyncResult.value` 屬性中的字串等形式提供。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|
####  <a name="setasync(location,-[options],-[callback])"></a>setAsync(location, [options], [callback])

設定約會的地點。

`setAsync` 方法會向 Exchange 伺服器啟動非同步呼叫，以設定約會的地點。設定約會地點會覆寫目前的地點。

##### <a name="parameters:"></a>參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`location`| 字串||約會的地點。字串限制在 255 個字元內。|
|`options`| 物件| &lt;選擇性&gt;|物件常值包含下列一個或多個屬性。<br/><br/>**屬性**<br/><table class="nested-table"><thead><tr><th>名稱</th><th>類型	</th><th>屬性</th><th>描述</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>物件</td><td>&lt;選擇性&gt;</td><td>開發人員可提供任何他們想要在回呼方法中存取的物件。</td></tr></tbody></table>|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 <br/>如果地點設定失敗，`asyncResult.error` 屬性將包含錯誤碼。<br/><table class="nested-table"><thead><tr><th>錯誤碼</th><th>描述</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td><code>location</code> 參數的長度超過 255 個字元。</td></tr></tbody></table>|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.1|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫|
