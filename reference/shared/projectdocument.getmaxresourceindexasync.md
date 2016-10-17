
# <a name="projectdocument.getmaxresourceindexasync-method-(javascript-api-for-office-v1.1)"></a>ProjectDocument.getMaxResourceIndexAsync 方法 (JavaScript API for Office v1.1)
以非同步方式取得目前專案中資源集合的最大索引。 **重要事項：**此 API 只適用於 Windows 桌面上的 Project 2016。

|||
|:-----|:-----|
|**主應用程式︰**|Project|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**已新增於**|1.1|

```js
Office.context.document.getMaxResourceIndexAsync([options][, callback]);
```


## <a name="parameters"></a>參數


_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;下列**[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)：**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;類型：**陣列**、**布林值**、**null**、**數字**、**物件**、**字串** 或 **未定義**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;無變更的情況下，於 [AsyncResult](../../reference/shared/asyncresult.md) 物件中傳回的任一類型使用者定義項目。選用。<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;例如，您可以傳遞 _asyncContext_ 引數，方法是使用格式 `{asyncContext: 'Some text'}` 或 `{asyncContext: <object>}`。

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**函數**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;方法呼叫傳回時所叫用的函數，其唯一的參數為 [AsyncResult](../../reference/shared/asyncresult.md) 類型。選用。
    

## <a name="callback-value"></a>回呼值

當 _callback_ 函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的參數存取該物件。

若為 **getMaxResourceIndexAsync** 方法，傳回的 [AsyncResult](../../reference/shared/asyncresult.md) 物件會包含下列屬性。



|**名稱**|**描述**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在選擇性 _asyncContext_ 參數中傳遞的資料 (如果有使用該參數)。|
|[error](../../reference/shared/asyncresult.error.md)|錯誤的相關資訊 (如果 **status** 屬性等於 **failed**)。|
|[status](../../reference/shared/asyncresult.status.md)|非同步呼叫的 **succeeded** 或 **failed** 狀態。|
|[value](../../reference/shared/asyncresult.value.md)|在目前專案的資源集合中，最大的索引編號。|

## <a name="remarks"></a>備註

您可以使用傳回的值搭配 [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) 方法，以取得資源 GUID。資源集合不包含位於 0 索引的資源。


## <a name="example"></a>範例

下列程式碼範例會呼叫 **getResourceTaskIndexAsync**，以取得目前專案中資源集合的最大索引。然後它會使用傳回的值和 [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) 方法，以取得每個資源 GUID。

此範例假設您的增益集參照 jQuery 程式庫，且下列頁面控制項定義在頁面內文的內容 div 中。




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var resourceGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getResourceInfo);
        });
    };

    // Get the maximum resource index, and then get the resource GUIDs.
    function getResourceInfo() {
        getMaxResourceIndex().then(
            function (data) {
                getResourceGuids(data);
            }
        );
    }

    // Get the maximum index of the resources for the current project.
    function getMaxResourceIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxResourceIndexAsync(
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

    // Get each resource GUID, and then display the GUIDs in the add-in.
    // There is no 0 index for resources, so start with index 1.
    function getResourceGuids(maxResourceIndex) {
        var defer = $.Deferred();
        for (var i = 1; i <= maxResourceIndex; i++) {
            getResourceGuid(i);
        }
        return defer.promise();
        function getResourceGuid(index) {
            Office.context.document.getResourceByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resourceGuids.push(result.value);
                        if (index == maxResourceIndex) {
                            defer.resolve();
                            $('#message').html(resourceGuids.toString());
                        }
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
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
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|已導入|

## <a name="see-also"></a>請參閱



#### <a name="other-resources"></a>其他資源


[getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)

[AsyncResult 物件](../../reference/shared/asyncresult.md)

[ProjectDocument 物件](../../reference/shared/projectdocument.projectdocument.md)
