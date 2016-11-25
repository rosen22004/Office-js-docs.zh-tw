

# <a name="office.useshortnamespace-method"></a>Office.useShortNamespace 方法
為完整的 `Office` 命名空間切換使用 `Microsoft.Office.WebExtension` 別名的開關。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```js
Office.useShortNamespace(useShortcut);
```


## <a name="parameters"></a>參數



_useShortcut_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**布林值**

    
&nbsp;&nbsp;&nbsp;&nbsp;**true** 為使用捷徑別名；**false** 為停用。預設值為 **true**。
    


## <a name="example"></a>範例



```js
function startUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(true);
    }
    else {
        Office.useShortNamespace(true);
    }
    write('Office alias is now ' + typeof Office);
}

function stopUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(false);
    }
    else {
        Office.useShortNamespace(false);
    }
    write('Office alias is now ' + typeof Office);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、Outlook、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增支援在 Access 內容增益集中呼叫這個方法。|
|1.0|已導入|