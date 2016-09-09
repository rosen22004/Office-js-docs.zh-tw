
# ProjectDocument.getSelectedResourceAsync 方法
以非同步方式在資源檢視中取得所選資源的 GUID。

|||
|:-----|:-----|
|**主機︰**|Project|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**已新增於**|1.0|

```
Office.context.document.getSelectedResourceAsync([options,] [callback]);
```


## 參數



|**名稱**|**類型	**|**說明**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _options_|**物件**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**物件**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## 回呼值

當 _callback_ 函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的參數存取該物件。

若為 **getSelectedResourceAsync** 方法，傳回的 [AsyncResult](../../reference/shared/asyncresult.md) 物件會包含下列屬性。


****


|**名稱**|**說明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在選擇性 _asyncContext_ 參數中傳遞的資料 (如果有使用該參數)。|
|[錯誤](../../reference/shared/asyncresult.error.md)|錯誤的相關資訊 (如果 **status** 屬性等於 **failed**)。|
|[狀態](../../reference/shared/asyncresult.status.md)|非同步呼叫的 **succeeded** 或 **failed** 狀態。|
|[value](../../reference/shared/asyncresult.value.md)|做為 **string** 的所選資源 GUID。|

## 備註

在 Project 增益集中，資源的 GUID 比資源名稱更好用。資源 GUID 可用於存取資源資訊，例如 Project Online 中的資源相關資料 (可透過用戶端物件模型 (CSOM) 取得)。您也可以將資源 GUID 儲存在區域變數中，並將它用於 [getResourceFieldAsync](../../reference/shared/projectdocument.gettaskasync.md) 方法。

如果使用中檢視不是資源檢視 (例如 [資源使用狀況] 或 [資源工作表] 檢視)，或如果沒有在資源檢視中選取任何資源，**getSelectedResourceAsync** 就會傳回 5001 錯誤 (內部錯誤)。請參閱 [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) 方法，以取得根據使用中檢視類型，使用 [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) 事件和 [getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md) 方法啟動按鈕的範例。


## 範例

下列程式碼範例會呼叫 **getSelectedResourceAsync**，以取得資源檢視中目前選取的資源 GUID。然後它會以遞迴方式呼叫 [getResourceFieldAsync](../../reference/shared/projectdocument.gettaskasync.md)，以取得三個資源欄位值。

此範例假設您的增益集參照 jQuery 程式庫，且已在頁面內文的內容 div 中定義下列頁面控制項。




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getResourceInfo);
        });
    };

    // Get the GUID of the resource and then get the resource fields.
    function getResourceInfo() {
        getResourceGuid().then(
            function (data) {
                getResourceFields(data);
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

    // Get the specified fields for the selected resource.
    function getResourceFields(resourceGuid) {
        var targetFields =
            [Office.ProjectResourceFields.Name, Office.ProjectResourceFields.Units, Office.ProjectResourceFields.BaseCalendar];
        var fieldValues = ['Name: ', 'Units: ', 'Base calendar: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == targetFields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }

            // If the call is successful, get the field value and then get the next field.
            else {
                Office.context.document.getResourceFieldAsync(
                    resourceGuid,
                    targetFields[index],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            fieldValues[index] += result.value.fieldValue;
                            getField(index++);
                        }
                        else {
                            onError(result.error);
                        }
                    }
                );
            }
        }
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**可用於需求集合**|Selection|
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|

## 請參閱



#### 其他資源


[getResourceFieldAsync 方法](../../reference/shared/projectdocument.getresourcefieldasync.md)

[AsyncResult 物件](../../reference/shared/asyncresult.md)

[ProjectDocument 物件](../../reference/shared/projectdocument.projectdocument.md)
