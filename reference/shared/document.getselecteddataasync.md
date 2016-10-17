
# <a name="document.getselecteddataasync-method"></a>Document.getSelectedDataAsync 方法
讀取文件目前的選取範圍中所包含的資料。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、PowerPoint、Project、Word|
|**可用於需求集合**|Selection|
|**上次變更於 Selection**|1.1|

```js
Office.context.document.getSelectedDataAsync(coercionType [, options], callback); 
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)<br/><table><tr><td></td><td><b>主應用程式支援</b></td></tr><tr><td><b>Office.CoercionType.Text</b> (字串)</td><td>僅限 Excel、Excel Online、PowerPoint、PowerPoint Online、Word 和 Word Online</td></tr><tr><td><b>Office.CoercionType.Matrix</b> (陣列的陣列)</td><td>僅限 Excel、Word 和 Word Online</td></tr><tr><td><b>Office.CoercionType.Table</b> ([TableData](../../reference/shared/tabledata.md) 物件)</td><td>僅限 Access、Excel、Word 和 Word Online</td></tr><tr><td><b>Office.CoercionType.Html</b></td><td>僅限 Word。</td></tr><tr><td><b>Office.CoercionType.Ooxml</b> (Office Open XML)</td><td>僅限 Word 和 Word Online</td></tr><tr><td><b>Office.CoercionType.SlideRange</b></td><td>僅限 PowerPoint 和 PowerPoint Online</td></tr></table>|要傳回的資料結構類型。必要。||
| _options_|**object**<br/><table><tr><td><i>valueFormat</i></td><td><b>[ValueFormat](../../reference/shared/valueformat-enumeration.md)</b></td><td>指定是否要傳回結果，包含其已格式化或未格式化的數字或日期值。</td><td></td></tr><tr><td><i>filterType</i></td><td>[FilterType](../../reference/shared/filtertype-enumeration.md)</td><td>指定擷取資料時是否要套用篩選。選用。</td><td>此參數在 Word 文件中會予以忽略。</td></tr><tr><td><i>asyncContext</i></td><td><b>array</b>、<b>boolean</b>、<b>null</b><b>number</b>,  <b>object</b>、<b>string</b> 或 <b>undefined</b></td><td>在 <b>AsyncResult</b> 物件中傳回之任何類型的使用者定義項目不會有任何改變。</td><td></td></tr></table>|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **getSelectedDataAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取目前選擇中的值，其會透過您使用 _coercionType_ 參數指定的資料結構或格式傳回。(如需有關資料強制型轉的詳細資訊，請參閱**註解**。)|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

在工作窗格或內容增益集，使用 **getSelectedDataAsync** 方法撰寫指令碼，以從使用者在文件、試算表、簡報或專案中的選擇讀取資料。例如，使用者在 Word 文件中選取內容之後，您可以使用 **getSelectedDataAsync** 方法以讀取該選擇，然後將其提交至 Web 服務，做為查詢或一些其他作業。

讀取此選擇之後，您也可以使用 [Document](../../reference/shared/document.setselecteddataasync.md) 物件的 [setSelectedDataAsync](../../reference/shared/document.addhandlerasync.md) 和 **addHandlerAsync** 方法，以[寫回至選擇或新增事件處理常式](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)，以偵測使用者是否變更選擇。

**getSelectedDataAsync** 方法可從選項中讀取 (只要選項為作用中)。在 Word 和 Excel 的增益集中，如果您需要建立永久的關聯性，以讀取並寫入使用者的選擇，而非使用 [Bindings.addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) 方法，以[繫結至該選擇](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。

使用 **getSelectedDataAsync** 方法的 _coercionType_ 參數，以指定所選取要讀取之資料的資料結構或格式。



|**指定的 _coercionType_**|**傳回的資料**|**Office 主應用程式支援**|
|:-----|:-----|:-----|
|**Office.CoercionType.Text** 或 `"text"`|字串。|Word、 Excel、 PowerPoint 和 Project。<br/><br/> **附註**：在 Excel 中，即使已選取儲存格的子集，仍會傳回整個儲存格內容。|
|**Office.CoercionType.Matrix** 或 `"matrix"`|陣列的陣列。例如，` [['a','b'], ['c','d']]` 適用於兩欄中兩列的選擇範圍。|Word 和 Excel。|
|**Office.CoercionType.Table** 或 `"table"`|[TableData](../../reference/shared/tabledata.md) 物件適用於讀取帶有標頭的表格。|Word 和 Excel。|
|**Office.CoercionType.Html** 或 `"html"`|以 HTML 格式。|僅限 Word。|
|**Office.CoercionType.Ooxml** 或 `"ooxml"`|以 Open Office XML (OpenXML) 格式。|僅限 Word。<br/><br/> **秘訣**：開發增益集程式碼時，您可以使用 _getSelectedDataAsync_ 方法的 `"ooxml"`**coercionType**，以查看您在 Word 文件中選取的內容如何定義為 OpenXML 標籤。然後，使用 [Document.setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) 方法之資料參數中的這些標籤，以該格式或結構，將內容寫入文件。例如，您可以 [將影像插入文件](http://blogs.msdn.com/b/officeapps/archive/2012/10/26/inserting-images-with-apps-for-office.aspx) 做為 OpenXML。|
|**Office.CoercionType.SlideRange** 或 "slideRange"|包含名稱為 “slide” (其包含所選取投影片的 ID、標題和索引) 之陣列的 JSON 物件。**附註：**若要選取多個投影片，使用者必須在**一般**、**大綱模式**或**投影片瀏覽**檢視中編輯簡報。此外，**母片檢視**不支援此方法。例如，`{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` 適用於兩張投影片的選擇範圍。|僅限 PowerPoint。|
如果選擇範圍的資料結構不符合指定的 _coercionType_，則 **getSelectedDataAsync** 方法將嘗試強制轉型資料至該類型或結構。如果選擇範圍無法強制轉型至您指定的 **Office.CoercionType**，則 **AsyncResult.status** 屬性會傳回 `"failed"`。


## <a name="example"></a>範例

若要讀取目前選擇範圍的值，您需要寫入讀取選擇範圍的回呼函數。下列範例顯示作法：


-  **傳遞匿名的回呼函數**，可將目前選擇範圍的值讀取至 _getSelectedDataAsync_ 方法的 **callback** 參數。
    
-  **讀取選取範圍**，以文字、未格式化，而且未篩選的方式。
    
-  在增益集頁面上**顯示值**。
    

```js
function getText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
        { valueFormat: "unformatted", filterType: "all" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            } 
            else {
                // Get selected data.
                var dataValue = asyncResult.value; 
                write('Selected data is ' + dataValue);
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


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|Selection|
|**最低權限等級**|[ReadDocument (需要有 ReadAllDocument 才可取得 Office Open XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1| 在 Word Online 中，新增支援 **Office.CoercionType.Matrix** 和 **Office.CoercionType.Table** 做為 _coercionType_ 參數。|
|1.1|在 iPad 版 Office 的 Excel、PowerPoint 和 Word 中，新增與 Windows 桌面版 Excel、PowerPoint 和 Word 相同的支援等級。|
|1.1| 在 Word Online 中，新增支援 **Office.CoercionType.Text** 做為 _coercionType_ 參數。|
|1.1|在 PowerPoint 的內容增益集中，您可以藉由傳遞  **Office.CoercionType.SlideRange** 做為 _getSelectedDataAsync_ 方法的 **coercionType** 參數，以取得投影片選取範圍的 ID、標題和索引。如需如何使用此值以導覽至目前選取的投影片，請參閱 [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) 方法主題。|
|1.0|已導入|
