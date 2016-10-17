
# <a name="asyncresult.asynccontext-property"></a>AsyncResult.asyncContext property
取得傳遞至叫用方法之選擇性 _asyncContext_ 參數的使用者定義項目，並保留傳遞時的狀態。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```
var myContext = asynchResult.asyncContext;
```


## <a name="return-value"></a>傳回值

傳回使用者定義項目 (可以是任何 JavaScript 類型︰**字串**、**數字**、**布林值**、**物件**、**陣列**、**Null** 或 **未定義**) 傳遞至叫用方法之選擇性 _asyncContext_ 參數。傳回**未定義**，如果您未傳遞任何項目至 _asyncContext_ 參數。


## <a name="example"></a>範例




```js
function getDataWithContext() {
    var format = "Your data: ";
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, { asyncContext: format }, showDataWithContext);
}

 function showDataWithContext(asyncResult) {
    write(asyncResult.asyncContext + asyncResult.value);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**||||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增對 Access 增益集的支援。|
|1.0|已導入|
