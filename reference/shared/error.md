
# Error 物件
提供在非同步資料作業期間發生之錯誤的特定資訊。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```
asyncResult.error
```


## 成員


**屬性**


|**名稱**|**說明**|
|:-----|:-----|
|[code](../../reference/shared/error.code.md)|取得錯誤的數字代碼。|
|[name](../../reference/shared/error.name.md)|取得錯誤的名稱。|
|[訊息](../../reference/shared/error.message.md)|取得錯誤的詳細描述。|

## 備註

**Error** 物件存取自 [AsyncResult](../../reference/shared/asyncresult.md) 物件，而 AsyncResult 物件是在做為非同步資料作業之_回呼_引數傳遞的函數中傳回，例如 [Document](../../reference/shared/document.setselecteddataasync.md) 物件的 **setSelectedDataAsync** 方法。


## 範例

以下範例使用 **setSelectedDataAsync** 方法，將選取的文字設定為 “Hello World!”，如果失敗，則會顯示 **Error** 物件的 **名稱** 和 **訊息** 屬性值。


```js
function setText() {

    Office.context.document.setSelectedDataAsync("Hello World!", {},
        function (asyncResult) {
            if (asyncResult.status === "failed")
            var err = asyncResult.error; 
                write(err.name + ": " + err.message);
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Mac 版 Outlook**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|是|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增對 Access 內容增益集的支援。|
|1.0|已導入|
