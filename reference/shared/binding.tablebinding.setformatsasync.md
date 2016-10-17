
# <a name="tablebinding.setformatsasync-method"></a>TableBinding.setFormatsAsync 方法
設定或更新指定項目和繫結表格中資料的格式設定。

|||
|:-----|:-----|
|**主應用程式︰**|Excel|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|不在集合中|
|**已新增於**|1.1|

```
bindingObj.setFormatsAsync(cellFormat [,options] , callback);
```


## <a name="parameters"></a>參數



|**名稱**|**類型**|**描述**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _cellFormat_|**array**|包含一或多個 JavaScript 物件的陣列，該物件指定要鎖定目標之儲存格與套用至這些儲存格的格式設定。必要。||
| _options_|**object**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**object**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## <a name="callback-value"></a>回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **goToByIdAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用於...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|設定格式時，因為沒有可擷取的資料或物件，所以一律傳回 **undefined**。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## <a name="remarks"></a>備註

 **指定 cellFormat 參數**

使用 _cellFormat_ 參數，以設定或變更儲存格格式設定值，例如寬度、高度、字型、背景、對齊方式等等。您傳遞做為 _cellFormat_ 參數的值，是包含一或多個 JavaScript 物件的**陣列**，該物件指定要鎖定目標的儲存格 ( `cells:`) 以及要套用的格式 ( `format:`)。

_cellFormat_ 陣列中的每個 JavaScript 物件具有此形式：

 `{cells:{`_cell_range_`}, format:{`_format_definition_`}}`

`cells:` 屬性可指定您希望使用下列其中一個值設定格式的範圍：


**儲存格屬性中的支援範圍**


|**儲存格範圍設定**|**描述**|
|:-----|:-----|
| `{row: i}`|指定延伸至表格中第 i 列資料的範圍。|
| `{column: i}`|指定延伸至表格中第 i 欄資料的範圍。|
| `{row: i, column: j}`|指定表格中從第 i 列到第 j 欄資料的儲存格範圍。|
| `Office.Table.All`|指定整個表格，包括欄標題、資料和總計 (若有的話)。|
| `Office.Table.Data`|僅指定表格中的資料 (無標頭和總計)。|
| `Office.Table.Headers`|僅指定標頭列。|


屬性可指定對應至 Excel 中 **[格式化儲存格]** 對話方塊中可用之設定子集的值 (按右鍵 `format:` **[格式化儲存格]** 或 **[首頁]**  >  **[格式化]**  >  **[格式化儲存格]**)。

您指定 `format:` 屬性的值，做為 JavaScript 物件常值中一或多個_屬性名稱_ - _值_組的清單。_屬性名稱_可指定要設定的格式設定屬性的名稱，而_值_可指定屬性值。您可以為特定格式指定多個值，例如字型色彩與大小。以下是三個 `format:` 屬性值範例：




```
//Set cells: font color to green and size to 15 points.
format: {fontColor : "green", fontSize : 15}
```




```
//Set cells: border to dotted blue.
format: {borderStyle: "dotted", borderColor: "blue"}
```




```
//Set cells: background to red and alignment to centered.
format: {backgroundColor: "red", alignHorizontal: "center"}
```

您可以藉由指定 `numberFormat:` 屬性中的數字格式 “code” 字串以指定數字格式。您可以指定的數字格式字串，對應至您可以在 Excel 設定的字串 (使用 **[格式化儲存格]** 對話方塊之 **[數字]** 索引標籤上的 **[自訂]** 類別)。此範例顯示如何將數字格式化成兩位小數的百分比：




```
format: {numberFormat:"0.00%"}
```

如需詳細資訊，請參閱如何[建立自訂的數字格式](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1#BM1)。



 **指定單一目標**

下列範例顯示 _cellFormat_ 值可將標題列的字型色彩設為紅色。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: Office.Table.Headers, format: {fontColor: "red"}}], 
    function (asyncResult){});
```

 **指定多個目標**

**SetFormatsAsync** 方法可支援在單一函數呼叫中，格式化繫結表格中的多個目標。若要這麼做，您會針對要格式化的每個目標，傳遞 _cellFormat_ 陣列中的物件清單。例如，下列程式碼行會將第一列的字型色彩設為黃色，並將第三列的第四個儲存格設為白色邊框和粗體字。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

若要在寫入資料時，在表格上設定格式，請使用 _Document.setSelectedDataAsync_ 或 _TableBinding.setDataAsync_ 方法的 [tableOptions](http://msdn.microsoft.com/library/4c1e13e9-b61a-47df-836c-3ca9aba4ca1c%28Office.15%29.aspx) 和 [cellFormat](http://msdn.microsoft.com/library/5b6ecf6f-c57f-4c0d-9605-59daee8fde13%28Office.15%29.aspx) 選擇性參數。

使用 **Document.setSelectedDataAsync** 和 **TableBinding.setDataAsync** 方法的選擇性參數設定格式，僅適用於第一次寫入資料時設定格式。若要在寫入資料之後變更格式，請使用下列方法：


- 若要更新儲存格格式設定，例如字型色彩和樣式，請使用**TableBinding.setFormatsAsync**方法 (此方法)。
    
- 若要更新表格選項，例如帶狀資料列和篩選按鈕，請使用 [TableBinding.setTableOptions](../../reference/shared/binding.tablebinding.settableoptionsasync.md) 方法。
    
- 若要清除格式設定，請使用 [TableBinding.clearFormats](../../reference/shared/binding.tablebinding.clearformatsasync.md) 方法。
    
 **Excel Online 的其他備註**

傳遞至 _cellFormat_ 參數之_格式化群組_的數目不可超過 100。單一格式化群組包括套用至儲存格指定範圍的一組格式。例如，下列呼叫會傳遞兩個格式化群組至 _cellFormat_。




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});

```

如需詳細資訊和範例，請參閱[如何格式化 Excel 增益集中的表格](../../docs/excel/format-tables-in-add-ins-for-excel.md)。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**||**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|不在集合中。|
|**最低權限等級**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 的支援。|
|1.1|已導入|
