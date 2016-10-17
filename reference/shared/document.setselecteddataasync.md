
# <a name="document.setselecteddataasync-method"></a>document.setSelectedDataAsync 方法
將資料寫入文件中目前的選取範圍。

|||
|:-----|:-----|
|**主應用程式︰**Access、Excel、PowerPoint、Project、Word、Word Online|**增益集類型：** 內容、工作窗格|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**上次變更於**|1.1|

```js
Office.context.document.setSelectedDataAsync(data [, options], callback(asyncResult));
```


## <a name="parameters"></a>參數

|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _data_|資料可以是下列任何資料型別：<ul><li><b>字串</b> (Office.CoercionType.Text) - 僅適用於 Excel、Excel Online、PowerPoint、PowerPoint Online、Word 和 Word Online。</li><li>陣列的<b>陣列</b> (Office.CoercionType.Matrix) - 僅適用於 Excel、Word 和 Word Online。</li><li>[TableData](../../reference/shared/tabledata.md) (Office.CoercionType.Table) - 僅適用於 Access、Excel、Word 和 Word Online</li><li><b>HTML</b>  (Office.CoercionType.Html) - 僅適用於 Word 和 Word Online。</li><li><b>Office Open XML</b>  (Office.CoercionType.Ooxml) - 僅適用於 Word 和 Word Online。</li><li><b>Base64 已編碼影像資料流</b>  (Office.CoercionType.Image) - 僅適用於 Excel、PowerPoint、Word 和 Word Online。</li></ul>|要在目前選取範圍中設定的資料。必要。|**上次變更於：**1.1。支援 Access 內容增益集需要 **Selection** 需求集合 1.1 或 更高版本。支援設定影像資料需要 **ImageCoercion** 需求集合 1.1 或更高版本。若要針對應用程式啟動進行此設定，請使用：<br/><br/>`<Requirements>`<br/>&nbsp;&nbsp;`<Sets DefaultMinVersion="1.1">`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`<Set Name="ImageCoercion"/>`<br/>&nbsp;&nbsp;`</Sets>`<br/>`</Requirements>`<br/><br/>ImageCoercion 功能的執行階段偵測可以利用下列程式碼執行︰<br/><br/>`if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {)) {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaImageCoercion();`<br/>`} else {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaOoxml();`<br/>`}`|
| _options_|**object**|指定一組[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。Options 物件可以包含下列屬性以設定選項︰<br/><ul><li>coercionType (<b><a href="735eaab6-5e31-4bc2-add5-9d378900a31b.htm">CoercionType </a></b>) - 指定如何強制轉型要設定的資料。如果未設定此選項，會設定 Office.CoercionType.Text 的預設值 coercionType。</li><li>tableOptions (<b>object</b> ) - 對於插入的表格，機碼值配對的清單，指定 <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">表格格式設定選項</a>，例如，標題列、總計列，以及帶狀資料列。 </li><li>cellFormat (<b>object</b> ) - 對於插入的表格，機碼值配對的清單，指定欄、列或儲存格的範圍，以及要套用至該範圍的<a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">儲存格格式設定</a>。 </li><li>imageLeft (<b>number</b> ) - 此選項適用於插入影像。表示與 PowerPoint 投影片左側相關聯的插入位置，及其與 Excel 中目前選取之儲存的關聯。Word 會忽略此值。此值以點為單位。</li><li>imageTop (<b>number</b> ) - 此選項適用於插入影像。表示與 PowerPoint 投影片頂端相關聯的插入位置，及其與 Excel 中目前選取之儲存的關聯。Word 會忽略此值。此值以點為單位。</li><li>imageWidth (<b>number</b> ) - 此選項適用於插入影像。表示影像寬度。如果已提供此選項，卻沒有 imageHeight，則影像將會縮放至符合影像寬度的值。如果已提供影像寬度與影像高度，則會視情況調整影像大小。如果未提供影像高度或寬度，則會使用預設的影像大小和外觀比例。此值以點為單位。</li><li>imageHeight (<b>number</b> ) - 此選項適用於插入影像。表示影像高度。如果已提供此選項，卻沒有 imageWidth，則影像將會縮放至符合影像高度的值。如果已提供影像寬度與影像高度，則會視情況調整影像大小。如果未提供影像高度或寬度，則會使用預設的影像大小和外觀比例。此值以點為單位。</li><li>asyncContext (<b>object \| value</b> ) - <a href="540c114f-0398-425c-baf3-7363f2f6bc47.htm">AsyncResult</a> 物件之 asyncContext 屬性上可用的使用者定義物件。當回呼是具名函數時，使用此物件以將物件或值提供給 <b>AsyncResult</b>。</li></ul>|_tableOptions_ 和 _cellFormat_ 選項已在 v1.1 新增，並且在 Excel 2013 和 Excel Online 中支援。<br/><br/>_imageLeft_ 和 _ImageTop_ 在 Excel 和 PowerPoint 中支援。|
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **setSelectedDataAsync** 方法的回呼函數中，[AsyncResult.value](../../reference/shared/asyncresult.value.md) 屬性會一律傳回 **undefined**，因為沒有可擷取的物件或資料。


## <a name="remarks"></a>備註

針對 _data_ 傳遞的值，包含要寫入目前選取範圍的資料。如果值為：


-  **字串：**將插入可強制轉型至**字串**的純文字或任何類型。
    
    
    
    在 Excel 中，您也可指定_資料_做為有效的公式，可將該公式新增至選取的儲存格。例如，將_資料_設定為 `"=SUM(A1:A5)"` 會加總指定範圍內的值。然而，當您在繫結儲存格上設定公式之後，則無法從繫結儲存格讀取新增的公式 (或任何預先存在的公式)。如果您在選取的儲存格上呼叫 [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) 方法以讀取其資料，該方法會僅傳回顯示在儲存格中的資料 (公式的結果)。
    
-  **陣列的陣列 ("matrix")：**將插入沒有標題的表格式資料。例如，若要將資料寫入兩個資料欄中的三個資料列，您可以傳遞陣列，如下所示：`[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`。若要寫入三個資料列中的單一資料欄，請傳遞陣列，如下所示：`[["R1C1"], ["R2C1"], ["R3C1"]]`
    
    
    
    在 Excel 中，您也可以指定_資料_做為陣列的陣列，其中包含有效的公式，可將該公式新增至選取的儲存格。例如，如果不會覆寫其他的資料，將_資料_設定為 `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]`會將這兩個公式新增至選取範圍。就像在單一儲存格上設定公式為 "text" 時，您無法在設定後讀取新增的公式 (或任何預先存在的公式) - 您僅可讀取公式的結果。
    
-  **[TableData](../../reference/shared/tabledata.md) 物件：**將插入帶有標題的表格。
    
    
    
     **附註︰**在 Excel 中，如果您在針對 _data_ 參數傳遞的 **TableData** 物件中指定公式，您可能不會得到預期的結果，因為 Excel 的「導出資料行」功能會自動複製資料欄中的公式。當您希望將包含公式的_資料_寫入選取的表格時，如果要解決此問題，請嘗試將資料指定為陣列的陣列 (而非 **TableData** 物件)，並將 _coercionType_ 指定為 **Microsoft.Office.Matrix** 或 “matrix”。
    
 **特定應用程式行為**

此外，將資料寫入選取範圍時，會套用這些應用程式特定的動作。

 **對於 Word**


- 如果沒有任何選取，而且插入點位於有效位置，則指定的 _data_ 會插入在如下所示的插入點：
    
      - 如果 _data_ 是字串，則插入指定的文字。
    
  - 如果 _data_ 是陣列的陣列 (“matrix”) 或 **TableData** 物件時，會插入新的 Word 表格。
    
  - 如果 _data_ 是 HTML，會插入指定的 HTML。
    
     >**重要事項**：如果您插入的任何 HTML 無效，Word 將不會顯示錯誤。Word 將會盡可能插入 HTML，因為它可以且會略過任何無效的資料。
  - 如果 _data_ 是 Office Open XML，會插入指定的 XML。
    
  - 如果 _data_ 是 base64 編碼的影像資料流，會插入指定的影像。
    
- 如果有選取範圍，則會以遵循上述相同規則的指定 _data_ 取代。
    
-  **插入影像**：插入的影像會放置內嵌。會忽略 **ImageLeft** 和 **imageTop** 參數。影像外觀比例永遠鎖定。如果僅提供 **imageWidth** 和 **imageHeight** 其中一個參數，其他值將會自動縮放以保持原始外觀比例。
    
 **對於 Excel**


- 如果選取單一儲存格：
    
      - 如果 _data_ 為字串，則指定的文字會插入為目前儲存格的值。
    
  - 如果 _data_ 是陣列的陣列 (“matrix”)，會插入指定的列和欄集合，如果周圍沒有其他資料，則會覆寫儲存格。
    
  - 如果 _data_ 是 **TableData** 物件，則會插入具有指定列和標頭集合的新 Excel 表格，如果周圍沒有其他資料，則會覆寫儲存格。
    
- 如果已選取多個儲存格，而且形狀不符合 _data_ 的形狀，則會傳回錯誤。
    
- 如果已選取多個儲存格，而且選取範圍的形狀完全符合 _data_ 的形狀，則會根據 _data_ 中的值更新所選取儲存格的值。
    
-  **插入影像**：插入的影像是浮動的。位置 **imageLeft** 和 **imageTop** 參數與目前選取的儲存格是相對的。Excel 允許 **imageLeft** 和 **imageTop** 負值，並且可能重新調整以在工作表內設定影像位置。影像外觀比例已鎖定，除非提供 **imageWidth** 和 **imageHeight** 參數。如果僅提供 **imageWidth** 和 **imageHeight** 其中一個參數，其他值將會自動縮放以保持原始外觀比例。
    
在其他情況下，會傳回錯誤。

 **對於 Excel Online**

除了 Excel 上述行為以外，在 Excel Online 中寫入資料時，會套用下列限制。 


- 您可以利用 _data_ 參數寫入工作表的儲存格總數，在針對此方法的單一呼叫中，不可超過 20,000。
    
- 傳遞至 _cellFormat_ 參數之_格式化群組_的數目不可超過 100。單一格式化群組包括套用至儲存格指定範圍的一組格式。例如，下列呼叫會傳遞兩個格式化群組至 _cellFormat_。
    

```js
  Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```

 **對於 PowerPoint**

插入的影像是浮動的。位置 **imageLeft** 和 **imageTop** 參數是選擇性的，但如果已提供，則兩者皆應該出現。如果提供單一值，則會忽略它。允許 **imageLeft** 和 **imageTop** 負值，並可將影像位置設定在投影片外。如果未提供任何選擇性參數，而且投影片有預留位置，則影像將取代投影片中的預留位置。影像外觀比例已鎖定，除非提供 **imageWidth** 和 **imageHeight** 參數。如果僅提供 **imageWidth** 和 **imageHeight** 其中一個參數，其他值將會自動縮放以保持原始外觀比例。


## <a name="example"></a>範例

下列範例將選取的文字或儲存格設定為"Hello World!"，如果失敗的話，則會顯示 [error.message](../../reference/shared/error.message.md) 屬性的值。


```js
function writeText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                 write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



指定選擇性的 _coercionType_ 參數，可讓您指定要寫入選取範圍的資料類型。下列範例會寫入資料做為兩欄三列的陣列，針對該資料結構，將 _coercionType_ 指定為 `"matrix"`，如果失敗，則會顯示 [error.message](../../reference/shared/error.message.md) 屬性的值。




```js
function writeMatrix() {
    Office.context.document.setSelectedDataAsync([["Red", "Rojo"], ["Green", "Verde"], ["Blue", "Azul"]], {coercionType: Office.CoercionType.Matrix}
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



下列範例會寫入具有一個標頭和四列的單欄表格，針對該資料結構，將 _coercionType_ 指定為 `"table"`，如果失敗，則會顯示 [error.message](../../reference/shared/error.message.md) 屬性的值。




```js
function writeTable() {
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [['Berlin'], ['Roma'], ['Tokyo'], ['Seattle']];

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: Office.CoercionType.Table},
        function (result) {
            var error = result.error
            if (result.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



 在 Word 中，如果您要寫入 HTML 至選取範圍，您可以將 _coercionType_ 參數指定為 `"html"`，如下列範例所示，其使用 HTML `<b>` 標籤，讓 “Hello” 變成粗體。




```js
function writeHtmlData() {
    Office.context.document.setSelectedDataAsync("<b>Hello</b> World!", {coercionType: Office.CoercionType.Html}, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在 Word、 PowerPoint 或 Excel 中，如果您想要撰寫影像到選取範圍中，您可以將 _coercionType_ 參數指定為 `"image"`，如下列範例所示。請注意，Word 會忽略 imageLeft 和 imageTop。




```js
function insertPictureAtSelection(base64EncodedImageStr) {

    Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
       coercionType: Office.CoercionType.Image,
       imageLeft: 50,
       imageTop: 50,
       imageWidth: 100,
       imageHeight: 100
       },
       function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log("Action failed with error: " + asyncResult.error.message);
           }
       });
}
```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的核取記號 (![Check symbol](../../images/mod_off15_checkmark.png))，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**

||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|![Check symbol](../../images/mod_off15_checkmark.png)|||
|**Excel**|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|
|**PowerPoint**|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|
|**Word**|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|![Check symbol](../../images/mod_off15_checkmark.png)|


|||
|:-----|:-----|
|**可用於需求集合**|Selection|
|**最低權限等級**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|在 Word 和 Word Online 中，新增支援寫入資料做為 base64 編碼的影像資料流。|
|1.1|在 Word Online 中，新增支援寫入 _data_ 做為**陣列的陣列** (matrix) 和 **TableData** (table)。|
|1.1|在 iPad 版 Office 的 Excel、PowerPoint 和 Word 中，新增與 Windows 桌面版 Excel、PowerPoint 和 Word 相同的支援等級。|
|1.1|在 Word Online 中，新增支援寫入 _data_ 做為**字串** (文字)。|
|1.1|新增支援使用 _tableOptions_ 和 _cellFormat_ 選擇性參數，利用 Excel 增益集[插入表格時設定格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)。|
|1.1|新增支援在 Access 增益集中寫入表格資料。|
|1.0|已導入|
