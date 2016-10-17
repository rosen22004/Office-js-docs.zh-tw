
# <a name="binding.getdataasync-method"></a>Binding.getDataAsync 方法
傳回包含在繫結的資料。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**上次變更於 TableBinding**|1.1|

```
bindingObj.getDataAsync([, options] , callback );
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|指定如何強制轉型要設定的資料。 ||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|指定叫用方法所傳回的值，例如數字和日期，是否以其套用的格式設定傳回。||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|指定擷取資料時，是否必須套用篩選。||
| _rows_|**Office.TableRange.ThisRow**| 指定預先定義的字串 “thisRow”，以取得目前選取列中的資料。|僅適用於 Access 之內容增益集中的表格繫結。|
| _startRow_|**number**|針對表格或矩陣繫結，為繫結中的資料子集指定以零為基礎的起始列。 ||
| _startColumn_|**number**|針對表格或矩陣繫結，為繫結中的資料子集指定以零為基礎的起始欄。 ||
| _rowCount_|**number**|針對表格或矩陣繫結，指定從 _startRow_ 偏移的列數。 ||
| _columnCount_|**number**|針對表格或矩陣繫結，指定從 _startColumn_ 偏移的欄數。||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **Binding.getDataAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取特定繫結中的值。如果已指定 _coercionType_ 參數 (而且呼叫已成功)，會以 [CoercionType](../../reference/shared/coerciontype-enumeration.md) 列舉主題中所述的格式傳回資料。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

如果省略選擇性參數，(當適用於資料的類型與格式時) 則會使用下列預設值。



|**參數**|**預設值**|
|:-----|:-----|
| _coercionType_|繫結之原始、未強制轉型的類型。|
| _valueFormat_|未格式化的資料。|
| _filterType_|(未篩選的) 所有值。|
| _startRow_|第一列。|
| _startColumn_|第一欄。|
| _rowCount_|所有的列。|
| _columnCount_|所有的欄。|
從 [MatrixBinding](../../reference/shared/binding.matrixbinding.md) 或 [TableBinding](../../reference/shared/binding.tablebinding.md) 呼叫時，如果已指定選擇性的**startRow**、_startColumn_、_rowCount_ 以及 _columnCount_ 參數 (而且指定連續且有效的範圍)，_getDataAsync_ 方法將傳回繫結值的子集。


## <a name="example"></a>範例




```
function showBindingData() {
    Office.select("bindings#MyBinding").getDataAsync(function (asyncResult) {
        write(asyncResult.value)
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



透過 _Binding.getDataAsync_ 方法使用 `"table"` 和 `"matrix"`**coercionType** 之間的行為，有一個重要的差異，關於以標題列格式化的資料，如下列兩個範例所示。這些程式碼範例顯示 [Binding.SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) 事件的事件處理常式函數。

如果您指定 `"table"` _coercionType_，則 [TableData.rows](../../reference/shared/tabledata.rows.md) 屬性 (在下列程式碼範例中的 `result.value.rows`) 會傳回僅包含表格主體資料列的陣列。因此，其第 0 列將會是表格中第一個非標題列。




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'table', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value.rows[0][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
}
```

不過，如果您指定 `"matrix"` _coercionType_，則下列程式碼範例中的 `result.value` 會傳回包含第 0 列中表格標題的陣列。如果表格標題包含多個列，則這些會在包含表格主體資料列之前，全部包含在 `result.value` 矩陣，做為個別列。




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'matrix', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value[1][0]); 
            } 
            else 
                write(result.error.message); 
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


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|MatrixBindings, TableBindings, TextBindings|
|**最低權限等級**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.1|新增對 Access 增益集中表格繫結的支援。|
|1.0|已導入|

## <a name="see-also"></a>請參閱



#### <a name="other-resources"></a>其他資源


[繫結至文件或試算表中的區域。](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
