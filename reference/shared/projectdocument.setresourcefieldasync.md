

# <a name="projectdocument.setresourcefieldasync-method"></a>ProjectDocument.setResourceFieldAsync 方法
以非同步方式設定指定資源的指定欄位值。 **重要事項：**此 API 只適用於 Windows 桌面上的 Project 2016。

|||
|:-----|:-----|
|**主應用程式︰**|Project|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**已新增於**|1.1|

```js
Office.context.document.setResourceFieldAsync(resourceId, fieldId, fieldValue[, options][, callback]);
```


## <a name="parameters"></a>參數

_resourceId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;資源的 GUID。必要。
    
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;目標欄位的 ID，做為 [ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md) 常數或其對應的整數值。必要。
    
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;目標欄位中的值為**字串**、**數字**、**布林值** 或 **物件**。必要。
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;下列[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)：

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;類型：**陣列、布林值、null、數字、物件、字串**或**未定義**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;無變更的情況下，於 [AsyncResult](../../reference/shared/asyncresult.md) 物件中傳回的任一類型使用者定義項目。選用。</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;例如，您可以傳遞 _asyncContext_ 引數，方法是使用格式 `{asyncContext: 'Some text'}` 或 `{asyncContext: <object>}`。


_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**函數**

&nbsp;&nbsp;&nbsp;&nbsp;方法呼叫傳回時所叫用的函數，其唯一的參數為 [AsyncResult](../../reference/shared/asyncresult.md) 類型。選用。

    

## <a name="callback-value"></a>回呼值

當 _callback_ 函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的參數存取該物件。

若為 **setResourceFieldAsync** 方法，傳回的 [AsyncResult](../../reference/shared/asyncresult.md) 物件會包含下列屬性。


|**名稱**|**描述**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在選擇性 _asyncContext_ 參數中傳遞的資料 (如果有使用該參數)。|
|[error](../../reference/shared/asyncresult.error.md)|錯誤的相關資訊 (如果 **status** 屬性等於 **failed**)。|
|[status](../../reference/shared/asyncresult.status.md)|非同步呼叫的 **succeeded** 或 **failed** 狀態。|
|[value](../../reference/shared/asyncresult.value.md)|這個方法不會傳回值。|

## <a name="remarks"></a>備註

請先呼叫 [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 或 [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) 方法，以取得資源 GUID，然後將此 GUID 當作 _resourceId_ 引數傳遞至 **setResourceFieldAsync**。在每個非同步呼叫中，只能更新單一資源的單一欄位。


## <a name="example"></a>範例

下列程式碼範例會呼叫 [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)，以取得資源檢視中目前選取的資源 GUID。然後它會以遞迴方式呼叫 **setResourceFieldAsync**，以取得兩個資源欄位值。

範例中使用的 [GetSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 方法需要工作檢視 (例如 [工作使用狀況]) 為使用中檢視，並已選取工作。請參閱 [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) 方法，以取得依據使用中檢視類型啟動按鈕的範例。

此範例假設您的增益集參照 jQuery 程式庫，且已在頁面內文的內容 div 中定義下列頁面控制項。




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setResourceInfo);
        });
    };

    // Get the GUID of the resource, and then get the resource fields.
    function setResourceInfo() {
        getResourceGuid().then(
            function (data) {
                setResourceFields(data);
            }
        );
    }

    // Get the GUID of the selected resource.
    function getResourceGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Set the specified fields for the selected resource.
    function setResourceFields(resourceGuid) {
        var targetFields = [Office.ProjectResourceFields.StandardRate, Office.ProjectResourceFields.Notes];
        var fieldValues = [.28, 'Notes for the resource.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setResourceFieldAsync(
                resourceGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
    }

    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**可用於需求集合**||
|**最低權限等級**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄

|**版本**|**變更**|
|:-----|:-----|
|1.1|已導入|

## <a name="see-also"></a>請參閱



#### <a name="other-resources"></a>其他資源


[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)
[getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)
[AsyncResult 物件](../../reference/shared/asyncresult.md)
[ProjectResourceFields 列舉](../../reference/shared/projectresourcefields-enumeration.md)
[ProjectDocument 物件](../../reference/shared/projectdocument.projectdocument.md)

