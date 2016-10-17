
# <a name="context.contentlanguage-property"></a>Context.contentLanguage 屬性
 取得使用者指定用於編輯文件或項目的地區設定 (語言)。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```
var myContentLang = Office.context.contentLanguage;
```


## <a name="return-value"></a>傳回值

RFC 1766 語言標記格式中的**字串**，例如 `en-US`。


## <a name="remarks"></a>備註

**contentLanguage** 值會反映透過 Office 主應用程式中的 **[檔案]**  >  **[選項]**  >  **[語言]** 指定的 **[編輯語言]** 設定。

在 Access web 應用程式的內容增益集中，**contentLanguage** 屬性會取得增益集的文化特性 (例如，"en-GB")。


## <a name="example"></a>範例




```js
function sayHelloWithContentLanguage() {
    var myContentLanguage = Office.context.contentLanguage;
    switch (myContentLanguage) {
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




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增在 Access 的內容增益集中存取這個 API 的方法。|
|1.0|已導入|
