
# <a name="asyncresult.value-property"></a>AsyncResult.value 屬性
如果有的話，請取得這個非同步作業的裝載或內容。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```js
var dataValue = asyncResult.value;
```


## <a name="return-value"></a>傳回值

在進行非同步呼叫時，傳回要求值。 


 >**附註**：**值**屬性針對特定「Async」方法傳回的值，根據該方法的用途與內容而有所不同。若要判斷**值**屬性針對「Async」方法傳回的內容，請參閱方法主題的「回呼值」一節。如需「Async」方法的完整清單，請參閱 [AsyncResult](../../reference/shared/asyncresult.md) 物件主題的「備註」一節。


## <a name="remarks"></a>備註

存取函數中的 **AsyncResult** 物件，以引數的方式傳遞至「Async」方法的 _callback_ 參數，例如 [Document](../../reference/shared/document.getselecteddataasync.md) 物件的 [getSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) 與 **setSelectedDataAsync** 方法。


## <a name="example"></a>範例




```js
function getData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Table, function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        }
        else {
            write(asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Office for Mac**|
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
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增對 Access 增益集的支援。|
|1.0|已導入|
