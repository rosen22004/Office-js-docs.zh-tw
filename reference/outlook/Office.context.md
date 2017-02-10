

# <a name="context"></a>內容

## <a name="officeofficemdcontext"></a>[Office](Office.md).context

Office.context 命名空間會提供共用的介面，可為所有 Office 應用程式中的增益集所使用。此清單會列出這些由 Outlook 增益集所使用的介面。有關 Office.context 命名空間的完整清單，請參閱 [在共用的 API 中的 Office.context 參考](../shared/office.context.md)。

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="namespaces"></a>命名空間

[信箱](Office.context.mailbox.md):提供用於存取 Microsoft Outlook 和網頁型 Outlook 的 Outlook 增益集物件模型。

### <a name="members"></a>成員

####  <a name="displaylanguage-string"></a>displayLanguage：字串

取得以 RFC 1766 語言標記格式的地區設定 (語言)，該設定是由使用者為 Office 主應用程式的 UI 所指定。

`displayLanguage` 值反映目前的 **[顯示語言]** 設定，其是在 Office 主應用程式中從 **[檔案] > [選項] > [語言]** 中所指定。

##### <a name="type"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a>officeTheme︰物件

提供 Office 佈景主題色彩屬性的存取。

> **附註：**iOS 版 Outlook 或 Android 版 Outlook 不支援這個成員。

使用 Office 佈景主題色彩，可讓您針對增益集的色彩配置以及使用者透過 **[檔案] > [Office 帳戶] > [Office 佈景主題] UI** 所選取的現行 Office 佈景主題 (套用於所有 Office 主應用程式)，進行協調。使用 Office 佈景主題色彩可適用於電子郵件與工作窗格增益集。

##### <a name="type"></a>類型：

*   物件

##### <a name="properties"></a>屬性：

|名稱| 類型	| 描述|
|---|---|---|
|`bodyBackgroundColor`| 字串|取得 Office 佈景主題的本文背景色彩作為十六進位色彩三元組。|
|`bodyForegroundColor`| String|取得 Office 佈景主題的本文前景色彩作為十六進位色彩三元組。|
|`controlBackgroundColor`| String|取得 Office 佈景主題的控制項背景色彩作為十六進位色彩三元組。|
|`controlForegroundColor`| 字串|取得 Office 佈景主題的本文控制項色彩作為十六進位色彩三元組。|

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.3|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsroamingsettingsmd"></a>roamingSettings :[RoamingSettings](RoamingSettings.md)

取得代表自訂設定的物件，或郵件增益集儲存至使用者信箱的狀態。

`RoamingSettings` 物件可讓您儲存並存取郵件增益集儲存於使用者信箱的資料，如此當從用來存取該信箱的任何主機用戶端應用程式上執行增益集時，便可讓其使用該資料。

##### <a name="type"></a>類型：

*   [RoamingSettings](RoamingSettings.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](./tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|
