
# <a name="projectdocument.getselecteddataasync-method"></a>ProjectDocument.getSelectedDataAsync 方法
以非同步方式取得甘特圖檢視中目前的選取範圍內，一或多個儲存格中所包含的資料文字值。

|||
|:-----|:-----|
|**主應用程式︰**|Project|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**已新增於**|1.0|

```
Office.context.document.getSelectedDataAsync(coercionType[, options][, callback]);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)|要傳回的資料結構類型。必要。<br/>Project 2013 只支援 **Office.CoercionType.Text** 或 `"text"`。||
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|要用於數字或日期值的格式。<br/>Project 2013 會忽略這個參數，並在內部將它設定為 `unformatted`。||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|指定只要包含可見資料或是所有資料。 <br/>Project 2013 會忽略這個參數，並在內部將它設定為  `all`。||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

當 _callback_ 函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的參數存取該物件。

若為 **getSelectedDataAsync** 方法，傳回的 [AsyncResult](../../reference/shared/asyncresult.md) 物件會包含下列屬性。


****


|**名稱**|**描述**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在選擇性 _asyncContext_ 參數中傳遞的資料 (如果有使用該參數)。|
|[error](../../reference/shared/asyncresult.error.md)|錯誤的相關資訊 (如果 **status** 屬性等於 **failed**)。|
|[status](../../reference/shared/asyncresult.status.md)|非同步呼叫的 **succeeded** 或 **failed** 狀態。|
|[value](../../reference/shared/asyncresult.value.md)|所選儲存格的文字值。|

## <a name="remarks"></a>備註

**ProjectDocument.getSelectedDataAsync** 方法會覆寫 [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) 方法，並傳回在甘特圖檢視中之一或多個儲存格中選取的資料文字值。**ProjectDocument.getSelectedDataAsync** 僅支援 [CoercionType](../../reference/shared/coerciontype-enumeration.md) 文字格式，不支援 `matrix`、`table` 或其他格式。


## <a name="example"></a>範例

下列程式碼範例會取得所選儲存格的值。它會使用選擇性 _asyncContext_ 參數，將一些文字傳遞至回呼函數。

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
            $('#get-info').click(getSelectedText);
        });
    };

    // Get the text from the selected cells in the document, and display it in the add-in.
    function getSelectedText() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            {asyncContext: 'Some related info'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'Selected text: {0}<br/>Passed info: {1}',
                        result.value, result.asyncContext);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
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
|**可用於需求集合**|Selection|
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Task pane|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|

## <a name="see-also"></a>請參閱



#### <a name="other-resources"></a>其他資源


[AsyncResult 物件](../../reference/shared/asyncresult.md)

[Office.CoercionType](../../reference/shared/coerciontype-enumeration.md)

[ProjectDocument 物件](../../reference/shared/projectdocument.projectdocument.md)
