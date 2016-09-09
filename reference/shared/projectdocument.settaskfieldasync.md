
# ProjectDocument.setTaskFieldAsync 方法 (JavaScript API for Office v1.1)
以非同步方式設定指定工作的指定欄位值。
 **重要事項：**此 API 只適用於 Windows 桌面上的 Project 2016。

|||
|:-----|:-----|
|**主應用程式︰**|Project|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**已新增於**|1.1|

```js
Office.context.document.setTaskFieldAsync(taskId, fieldId, fieldValue[, options][, callback]);
```


## 參數


_taskId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;任務的 GUID。 必要。<br/><br/>
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;目標欄位的 ID，做為 [ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md) 常數或其對應的整數值。 必要。<br/><br/>
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;目標欄位中的值為**字串**、**數字**、**布林值** 或 **物件**。 必要。<br/><br/>
__<br/>
&nbsp;&nbsp;&nbsp;&nbsp;下列[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)：<br/><br/>

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;類型：**陣列、布林值、null、數字、物件、字串**或**未定義**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;無變更的情況下，於 [AsyncResult](../../reference/shared/asyncresult.md) 物件中傳回的任一類型使用者定義項目。 選用。</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;例如，您可以傳遞 _asyncContext_ 引數，方法是使用格式 `{asyncContext: 'Some text'}` 或 `{asyncContext: <object>}`。<br/><br/>
_回呼_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**函數**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;方法呼叫傳回時所叫用的函數，其唯一的參數為 [AsyncResult](../../reference/shared/asyncresult.md) 類型。 選用。
    

## 回呼值

當 _callback_ 函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的參數存取該物件。

若為 **setTaskFieldAsync** 方法，傳回的 [AsyncResult](../../reference/shared/asyncresult.md) 物件會包含下列屬性。



|**名稱**|**說明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在選擇性 _asyncContext_ 參數中傳遞的資料 (如果有使用該參數)。|
|[錯誤](../../reference/shared/asyncresult.error.md)|錯誤的相關資訊 (如果 **status** 屬性等於 **failed**)。|
|[狀態](../../reference/shared/asyncresult.status.md)|非同步呼叫的 **succeeded** 或 **failed** 狀態。|
|[值](../../reference/shared/asyncresult.value.md)|這個方法不會傳回值。|

## 備註

請先呼叫 [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) 或 [getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md) 方法，以取得工作 GUID，然後將此 GUID 當作 _taskId_ 引數傳遞至 **setTaskFieldAsync**。在每個非同步呼叫中，只能更新單一工作的單一欄位。


## 範例

下列程式碼範例會呼叫 [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)，以取得工作檢視中目前所選工作的 GUID。然後它會以遞迴方式呼叫 **setTaskFieldAsync**，以設定兩個工作欄位值。

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
            $('#set-info').click(setTaskInfo);
        });
    };

    // Get the GUID of the task, and then get the task fields.
    function setTaskInfo() {
        getTaskGuid().then(
            function (data) {
                setTaskFields(data);
            }
        );
    }

    // Get the GUID of the selected task.
    function getTaskGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedTaskAsync(
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

    // Set the specified fields for the selected task.
    function setTaskFields(taskGuid) {
        var targetFields = [Office.ProjectTaskFields.Active, Office.ProjectTaskFields.Notes];
        var fieldValues = [true, 'Notes for the task.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setTaskFieldAsync(
                taskGuid,
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


## 支援詳細資料


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

## 支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|已導入|

## 請參閱



#### 其他資源


[getSelectedTaskAsync 方法](../../reference/shared/projectdocument.getselectedresourceasync.md)
[getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md)
[AsyncResult 物件](../../reference/shared/asyncresult.md)
[ProjectTaskFields 列舉](../../reference/shared/projecttaskfields-enumeration.md)
[ProjectDocument 物件](../../reference/shared/projectdocument.projectdocument.md)
