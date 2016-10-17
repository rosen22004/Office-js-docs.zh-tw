

# <a name="context"></a>內容

## [Office](Office.md). context

Office.context 命名空間會提供共用的介面，可為所有 Office 應用程式中的增益集所使用。此清單會列出這些由 Outlook 增益集所使用的介面。有關 Office.context 命名空間的完整清單，請參閱 [在共用的 API 中的 Office.context 參考](../../shared/office.context.md)。


##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="namespaces"></a>命名空間

[信箱](Office.context.mailbox.md):提供用於存取 Microsoft Outlook 和網頁型 Outlook 的 Outlook 增益集物件模型。

### <a name="members"></a>成員

####  <a name="displaylanguage-:string"></a>displayLanguage：字串

取得以 RFC 1766 語言標記格式的地區設定 (語言)，該設定是由使用者為 Office 主應用程式的 UI 所指定。

`displayLanguage` 值反映目前的 **[顯示語言]** 設定，其是在 Office 主應用程式中從 **[檔案] > [選項] > [語言]** 中所指定。

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
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

####  <a name="roamingsettings-:[roamingsettings](roamingsettings.md)"></a>roamingSettings :[RoamingSettings](RoamingSettings.md)

取得代表自訂設定的物件，或郵件增益集儲存至使用者信箱的狀態。

`RoamingSettings` 物件可讓您儲存並存取郵件增益集儲存於使用者信箱的資料，如此當從用來存取該信箱的任何主機用戶端應用程式上執行增益集時，便可讓其使用該資料。

##### <a name="type:"></a>類型：

*   [RoamingSettings](RoamingSettings.md)

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| 限制|
|適用的 Outlook 模式| 撰寫或讀取|
