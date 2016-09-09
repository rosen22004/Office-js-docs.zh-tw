

# RoamingSettings

使用 `RoamingSettings` 物件的方法所建立的設定，會按照每個增益集和每位使用者來儲存。 也就是說，它們只能用於建立它們的增益集，而且只能從儲存它們的使用者信箱取用。

> 雖然 Outlook 增益集 API 限制對這些設定的存取，僅限建立它們的增益集，不過仍不應將這些設定視為安全儲存。可由 Exchange Web Services 或 Extended MAPI 來存取這些設定。不應該將它們用來儲存如使用者認證或安全性權杖之類的敏感資訊。

設定的名稱為字串，而值可以是字串、數字、布林值、null、物件或陣列。

`RoamingSettings` 物件可透過 `Office.context` 命名空間中的 [`roamingSettings`](Office.context.md#roamingsettings-roamingsettings) 屬性存取。

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

### 範例

```
// Get the current value of the 'myKey' setting
var value = Office.context.roamingSettings.get('myKey');
// Update the value of the 'myKey' setting
Office.context.roamingSettings.set('myKey', 'Hello World!');
// Persist the change
Office.context.roamingSettings.saveAsync();
```

### 方法

####  get(name) → (可為 null) {字串 | 數值 | 布林值 | 物件 | 陣列}

擷取指定的設定。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`name`| String|要擷取之設定的區分大小寫名稱。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|

##### 傳回：

<dl class="param-type">

<dt>類型</dt>

<dd>字串 | 數值 | 布林值 | 物件 | 陣列</dd>

</dl>

####  remove(name)

移除指定的設定。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`name`| String|要移除之設定的區分大小寫名稱。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|
####  saveAsync([callback])

儲存設定。

增益集先前儲存的所有設定會在它初始化時載入，所以在工作階段的存留期間內，您只要使用 [`set`](RoamingSettings.md#setname-value) 和 [`get`](RoamingSettings.md#getname--nullable-stringnumberbooleanobjectarray) 方法就能使用設定屬性包在記憶體內的複本。 如果您想要保存設定，以便在下次使用增益集時使用，請使用 `saveAsync` 方法。

##### 參數：

|名稱| 類型	| 屬性| 描述|
|---|---|---|---|
|`callback`| 函數| &lt;選擇性&gt;|當方法完成時，在 `callback` 參數中傳遞的函數會以單一參數 `asyncResult`，也就是 [`AsyncResult`](simple-types.md#asyncresult) 物件進行呼叫。 |

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|
####  set(name, value)

設定或建立指定的設定。

set 方法會建立具有指定名稱的新設定 (如果設定不存在)，或設定具有指定名稱的現有設定。值會儲存在文件中，成為其資料類型的序列化 JSON 表示法。

每個增益集最多可使用 2 MB 的設定，每項個別設定僅限 32 KB。

除非呼叫 [`saveAsync`](RoamingSettings.md#saveasynccallback) 函數，否則任何使用 `set` 函數對設定所做的變更將不會儲存到伺服器。

##### 參數：

|名稱| 類型	| 描述|
|---|---|---|
|`name`| String|要設定或建立設定的區分大小寫名稱。|
|`value`| 字串 &#124; 數值 &#124; 布林值 &#124; 物件 &#124; 陣列|要儲存的值。|

##### 需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|
