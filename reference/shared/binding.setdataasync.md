
# <a name="binding.setdataasync-method"></a>Binding.setDataAsync 方法
將資料寫入指定的繫結物件所代表文件的繫結區段。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**上次變更於 TableBinding**|1.1|

```js
bindingObj.setDataAsync(data [, options] ,callback);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _data_|<table><tr><td><b>string</b></td><td>僅 Excel、Excel Online、Word 和 Word Online</td></tr><tr><td><b>array</b> (陣列的陣列 – “matrix”)</td><td>僅 Excel 和  Word</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp161002">
  <b>TableData</b></a></td><td>僅 Access、Excel 和 Word</td></tr><tr><td><b>HTML</b></td><td>僅限 Word 和 Word Online</td></tr><tr><td><b>Office Open XML</b></td><td>僅 Word</td></tr></table>|要在目前選取範圍中設定的資料。必要。|**上次變更於：**1.1 支援 Access 的內容增益集需要 **TableBinding** 需求集合 1.1 或更高版本。|
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|指定如何強制轉型要設定的資料。 ||
| _columns_|**字串陣列**| 指定資料欄名稱。|**已新增於：**v1.1.僅適用於 Access 之內容增益集中的表格繫結。|
| _rows_|**Office.TableRange.ThisRow**|指定預先定義的字串 “thisRow”，以設定目前選取列中的資料。 |**已新增於：**v1.1.僅適用於 Access 之內容增益集中的表格繫結。|
| _startColumn_|**number**|針對資料的子集指定以零為基礎的起始欄。 |僅適用於表格或矩陣繫結。如果省略，資料會設定在第一欄開始。|
| _startRow_|**number**|針對繫結中資料的子集，指定以零為基礎的起始列。 |僅適用於表格或矩陣繫結。如果省略，資料會設定在第一列開始。|
| _tableOptions_|**object**|對於插入的表格，指定[表格格式化選項](../../docs/excel/format-tables-in-add-ins-for-excel.md)的機碼值組清單，例如標題列、總計列，以及帶狀資料列。 |**已新增於：** v1.1。**可支援：**Excel。|
| _cellFormat_|**object**|對於插入的表格，指定資料欄、資料列之範圍的機碼值組清單，或套用該範圍的儲存格及[儲存格格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)。|**已新增於** v1.1。**可支援：**Excel、Excel Online。|
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **setDataAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|因為沒有可擷取的物件或資料，所以一律傳回 **undefined**。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

針對_資料_傳遞的值包含要在繫結中寫入的資料。傳遞的值類型可決定要寫入的內容，如下表中所述。



|**_data_值**|**寫入資料**|
|:-----|:-----|
|**字串**|將寫入可強制轉型至**字串**的純文字或任何類型。|
|陣列的陣列 (“matrix”)|將寫入沒有標題的表格式資料。例如，若要將資料寫入兩個資料欄中的三個資料列，您可以傳遞陣列，如下所示：` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`若要寫入三個資料列中的單一資料欄，請傳遞陣列，如下所示：`[["R1C1"], ["R2C1"], ["R3C1"]]`|
|[TableData](../../reference/shared/tabledata.md) 物件|將寫入帶有標題的表格。|
此外，將資料寫入繫結時，會套用這些應用程式特定的動作。

 **對於 Word**，指定的_資料_已寫入繫結，如下所示：



|**_data_值**|**寫入資料**|
|:-----|:-----|
|**字串**|已寫入指定的文字。|
|陣列的陣列 (“matrix”) 或 **TableData** 物件|已寫入 Word 表格。|
|HTML|已寫入指定的 HTML。
 >**重要**  如果您寫入的任何 HTML 無效，Word 將不會顯示錯誤。Word 將會盡可能寫入 HTML，而且將省略任何無效資料。將會盡量 HTML 撰寫，以及它可以將略過任何無效的資料。

|
|Office Open XML ("Open XML")|已寫入指定的 XML。|  **對於 Excel**，指定的_資料_已寫入繫結，如下所示：



|**_data_值**|**寫入資料**|
|:-----|:-----|
|**字串**|指定的文字已插入做為第一個繫結儲存格的值。您也可以指定有效的公式，以將該公式新增至繫結儲存格。例如，將_資料_設定為 `"=SUM(A1:A5)"` 會加總指定範圍內的值。然而，當您在繫結儲存格上設定公式之後，則無法從繫結儲存格讀取新增的公式 (或任何預先存在的公式)。如果您在繫結儲存格上呼叫 [Binding.getDataAsync](../../reference/shared/binding.getdataasync.md) 方法以讀取其資料，該方法會僅傳回顯示在儲存格中的資料 (公式的結果)。|
|陣列的陣列 (“matrix”)，以及圖形完全符合指定繫結的圖形|資料列與資料欄的集合已寫入。您也可以指定包含有效公式之陣列的陣列，以將其新增至繫結儲存格。例如，將_資料_設定為 `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]`，將新增這兩個公式到包含兩個儲存格的繫結。就像在單一繫結儲存格上設定公式時，您無法使用 **Binding.getDataAsync** 方法從繫結讀取新增的公式 (或任何預先存在的公式) - 它只會傳回繫結儲存格中顯示的資料。|
|**TableData** 物件，以及檔案形狀符合繫結表格。|如果周圍儲存格中未覆寫任何其他資料，則會寫入指定的資料列集合及/或標題。**附註：**如果您在針對 **data** 參數傳遞的 _TableData_ 物件中指定公式，您可能不會得到預期的結果，因為 Excel 的「導出資料行」功能會自動複製資料欄中的公式。當您希望將包含公式的_資料_寫入繫結表格時，如果要解決此問題，請嘗試將資料指定為陣列的陣列 (而非 **TableData** 物件)，並將 _coercionType_ 指定為 **Microsoft.Office.Matrix** 或 “matrix”。|
 **Excel Online 的其他備註**


- 要傳遞給 _data_ 參數的值中的儲存格總數，在針對此方法的單一呼叫中，不可超過 20,000。
    
- 傳遞至 _cellFormat_ 參數之_格式化群組_的數目不可超過 100。單一格式化群組包括套用至儲存格指定範圍的一組格式。例如，下列呼叫會傳遞兩個格式化群組至 _cellFormat_。
    
```js
  Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});

```

在其他情況下，會傳回錯誤。

如果已指定選擇性的 **startRow** 和 _startColumn_ 參數，則 _setDataAsync_ 方法將寫入表格子集或矩陣繫結中的資料，而且它們會指定有效範圍。


## <a name="example"></a>範例




```js
function setBindingData() {
    Office.select("bindings#MyBinding").setDataAsync('Hello World!', function (asyncResult) { });
}
```

指定選擇性的 _coercionType_ 參數，可讓您指定要寫入繫結的資料類型。例如，在 Word 中，如果您要寫入 HTML 至文字繫結，您可以將 _coercionType_ 參數指定為 `"html"`，如下列範例所示，其使用 HTML `<b>` 標籤，讓 “Hello” 變成粗體。




```js
function writeHtmlData() {
    Office.select("bindings#myBinding").setDataAsync("<b>Hello</b> World!", {coercionType: "html"}, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在此範例中，呼叫 **setDataAsync** 會傳遞 _data_ 參數做為陣列的陣列 (以建立三個資料列的單一行)，並將帶有 _coercionType_ 參數的資料結構指定為 `"matrix"`。




```js
function writeBoundDataMatrix() {
    Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],{ coercionType: "matrix" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在此範例的 `writeBoundDataTable` 函數中，呼叫 **setDataAsync** 會傳遞 _data_ 參數做為 **TableData** 物件 (以寫入三個資料列的三個資料欄)，並將帶有 _coercionType_ 參數的資料結構指定為 `"table"`。 

在 `updateTableData` 函數中，再次呼叫 **setDataAsync** 會傳遞 _data_ 參數做為 **TableData** 物件，但會做為具有新標頭和三個資料列的單一資料欄，以更新在 `writeBoundDataTable` 函數中建立之表格的最後一欄中的值。以零為基礎的選擇性 _startColumn_ 參數已指定為 2，以取代表格之第三欄中的值。




```js
function writeBoundDataTable() {
    // Create a TableData object.
    var myTable = new Office.TableData();
    myTable.headers = ['First Name', 'Last Name', 'Grade'];
    myTable.rows = [['Kim', 'Abercrombie', 'A'], ['Junmin','Hao', 'C'],['Toni','Poe','B']];

    // Set myTable in the binding.
    Office.select("bindings#myBinding").setDataAsync(myTable, { coercionType: "table" }, 
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Error: '+ asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}

// Replace last column with different data.
function updateTableData() {
     var newTable = new Office.TableData();
     newTable.headers = ["Gender"];
     newTable.rows = [["M"],["M"],["F"]];
     Office.select("bindings#myBinding").setDataAsync(newTable, { coercionType: "table", startColumn:2 }, 
         function (asyncResult) {
             if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                 write('Error: '+ asyncResult.error.message);
         } else {
            write('Bound data: ' + asyncResult.value);
         }     
     });   
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
|**最低權限等級**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援。|
|1.1|<ul><li>在 Access 的增益集中，新增支援寫入表格資料。</li><li>在 Excel 的增益集中，新增支援使用 <span class="parameter" sdata="paramReference">tableOptions</span> 和 <span class="parameter" sdata="paramReference">cellFormat</span> 選擇性參數<a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">寫入資料至表格繫結時的設定格式</a>。</li></ul>|
|1.0|已導入|

## <a name="see-also"></a>請參閱



#### <a name="other-resources"></a>其他資源


[繫結至文件或試算表中的區域。](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
