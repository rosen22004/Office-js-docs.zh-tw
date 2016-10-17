
# <a name="format-tables-in-add-ins-for-excel"></a>格式化 Excel 增益集中的表格


本文說明格式設定 API 的不同功能，並概述如何使用它們。在這個版本中，以程式設計方式指定儲存格格式設定與一些其他選項僅適用於表格 (不適用 **Office.CoercionType.Text** 或 **Office.CoercionType.Matrix** 資料結構)，而且僅適用於 Excel 增益集。若要使用增益集設定格式設定︰

- 使用者選取表格 (或以程式設計的方式插入表格的位置)，然後增益集可以在該表格上呼叫 **Document.setSelectedDataAsync** 方法來設定格式設定。

- 或者，如果活頁簿已包含繫結表格 (或增益集使用 [Bindings](../../reference/shared/bindings.bindings.md) 物件的一個 "addFrom" 方法在初始化時建立繫結表格)，增益集可以在那些繫結表格上呼叫 **Binding.setDataAsync** 方法來設定格式設定。
    
>**重要事項：**若要使用這些新的及更新的方法來格式化在 Excel 增益集中的表格，您的增益集專案必須[使用或更新為使用 Office.js 1.1 版或更新版本](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。

## <a name="specifying-formatting"></a>指定格式設定

若要指定您想要設定的格式設定，您可以建立包含一或多個機碼值組的 JavaScript 物件常值。您可以在 JavaScript 物件內，於清單中結合一系列的格式設定。例如： 


```js
var myFormat = {fontStyle:"bold", width:"autoFit", borderColor:"purple"};
```

若要套用格式設定，請將 JavaScript 物件傳遞至支援格式設定資料和表格其他功能的其中一個方法。

您可以透過兩種方式使用格式設定︰


- 增益集第一次將資料寫入至選取範圍或繫結時，藉由在傳遞至 [Document.setSelectedDataAysnc](../../reference/shared/document.setselecteddataasync.md) 或 [Binding.setDataAsync](../../reference/shared/binding.setdataasync.md) 方法的 _options_ 物件中指定選擇性的 _cellFormat_ 或 _tableOptions_ 參數。
    
- 最初將格式設定之後，您可以使用此用途專屬的新方法之一來[清除或更新格式設定](#updating-and-clearing-formatting)。
    

## <a name="using-optional-parameters-with-data-setting-methods"></a>使用選擇性參數搭配資料設定方法

針對表格繫結，您可以使用下列選擇性參數，在使用 **Document.setSelectedData** 或 **Binding.setDataAsync** 方法設定資料時指定格式設定︰_tableOptions_ 和 _cellFormat_。


### <a name="the-tableoptions-optional-parameter"></a>tableOptions 選擇性參數

使用 _tableOptions_ 選擇性參數來指定預設的表格樣式，並開啟及關閉某些表格的功能，例如︰**標頭列**、**合計列**和**帶狀列**。您傳遞做為 _tableOptions_ 參數的值是包含一份機碼值組的 JavaScript 物件。例如，


```js
tableOptions: {bandedRows: true, filterButton: false, style:"TableStyleMedium3"};
```


### <a name="the-cellformat-optional-parameter"></a>cellFormat 選擇性參數

使用 _cellFormat_ 選擇性參數，以變更儲存格格式設定值，例如寬度、高度、字型、背景、對齊方式等等。您傳遞做為 _cellFormat_ 參數的值，是包含 JavaScript 物件的陣列，該物件指定要鎖定目標的儲存格以及要套用的格式。例如：


```js
cellFormat: 
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: Office.Table.Headers, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}]
```

您可以在 _cellFormat_ 陣列中結合多個 `cells:` 和 `format:` 配對，以減少套用格式設定所需的函式呼叫的數目。


#### <a name="cells"></a>儲存格

使用 `cells:` 來指定您要套用格式欄、列和儲存格的範圍。


**儲存格值中的支援範圍**


|**儲存格範圍設定**|**描述**|
|:-----|:-----|
| `{row: i}`|指定延伸至表格中第 i 列資料的範圍。|
| `{column: i}`|指定延伸至表格中第 i 欄資料的範圍。|
| `{row: i, column: j}`|指定表格中從第 i 列到第 j 欄資料的儲存格範圍。|
| `Office.Table.All`|指定整個表格，包括欄標題、資料和總計 (若有的話)。|
| `Office.Table.Data`|僅指定表格中的資料 (無標頭和總計)。|
| `Office.Table.Headers`|僅指定標頭列。|

#### <a name="format"></a>格式

使用 `format:` 來指定您要套用到使用 `cells:` 定義為 JavaScript 機碼值組清單之範圍的格式設定。如需支援的值清單，請參閱[支援的格式設定機碼和值](#supported-formatting-keys-and-values)。

 **指定 Excel Online 格式設定的限制**

在 Excel Online 中設定格式設定時，傳遞至 _cellFormat_ 參數之_格式設定群組_的數目不可超過 100。單一格式設定群組包括套用至指定之儲存格範圍的一組格式設定。(換句話說，在陣列中其中一個 `cells:` 物件常值中指定的各個項目都會傳遞至 _cellFormat_。)例如，下列呼叫會傳遞兩個格式設定群組至 _cellFormat_。




```js
Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```


#### <a name="applying-optional-parameters"></a>套用選擇性的參數

在這個版本中，只有 **Document.setSelectedDataAsync** 和 **TableBinding.setDataAsync** 方法支援在相同的呼叫中使用 _tableOptions_ 和 _cellFormat_ 選擇性參數對表格寫入資料和設定格式設定。在下列範例中，傳遞給每個方法的第一個參數的 `tableData` 值 (_data_ 參數)，必須是包含要寫入之表格及資料定義的 [TableData](../../reference/shared/tabledata.md) 物件。

 **Document.setSelectedDataAsync 範例**




```js
Office.context.document.setSelectedDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 **TableBinding.setDataAsync 範例**




```js
Office.select("bindings#myBinding").setDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 >**附註：**：對 `Office.select("bindings#myBinding")` 的呼叫假設名為 `myBinding` 的繫結已經存在於工作表中。


## <a name="updating-and-clearing-formatting"></a>更新和清除格式設定


使用 **Document.setSelectedDataAsync** 或 **TableBinding.setDataAsync** 方法的 _cellFormat_ 和 _tableOptions_ 選擇性參數來設定格式設定時，它們只會在您第一次呼叫它們時設定格式設定。若要更新或清除格式設定，您必須使用 **TableBinding** 物件的三個新方法︰**setFormatsAsync**、**setTableOptionsAsync** 和 **clearFormatsAsync**。


### <a name="updating-formatting"></a>更新格式設定

[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md) 方法僅用於更新儲存格格式設定，例如寬度、高度、字型、背景和對齊方式。它會取得 _cellFormat_ 做為必要參數︰


```js
Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

[TableBinding.setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md) 方法僅用於更新表格選項，例如帶狀資料列和篩選按鈕。它會取得 _tableOptions_ 做為必要參數︰




```js
var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
```


### <a name="clearing-formatting"></a>清除格式設定

[TableBinding.clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md) 方法用於清除表格中所有格式設定。它會取得 _asyncContext_ 選擇性參數，以及選擇性的回撥函式︰


```js
Office.select("bindings#myBinding").clearFormatsAsync();
```


## <a name="supported-formatting-keys-and-values"></a>支援的格式設定機碼和值


下表列出您可以傳遞至 _cellFormat_ 或 _tableOptions_ 參數所支援的機碼值組。

針對 `format:` 值，可用的設定會與 [儲存格格式] 對話方塊中這些設定的子集合對應 (在功能區 [首頁] 索引標籤上，按一下滑鼠右鍵 > [儲存格格式] 或 [格式] > ** > [儲存格格式]**)。針對 `tableOptions:` 值，設定會與功能區上 [表格工具] |[設計] 索引標籤上的 [表格樣式選項] 和 [表格樣式] 群組中的設定對應。


 >**重要事項**：格式設定 API 的方法僅支援以下摘要的選項和值。如果您指定上述以外的格式設定選項或值時，處理行為會是未定義。這些未定義的處理行為在支援的平台之間不一定一致；您不應該為任何特定平台根據任何這些未定義的行為產生的副作用開發您的增益集。不過，未定義的處理行為不應該危害增益集狀態和 UI 或與它們互動的文件。


**對齊**


|**索引鍵**|**值**|**附註**|
|:-----|:-----|:-----|
| `alignHorizontal:`|"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "center across selection" \| "distributed"|當與縮排值組合時，僅支援下列的組合︰<br/><br/><ul><li><code>alignHorizontal: "left"</code> 和 <code>indentLeft: \<value\></code></li></ul><ul><li><code>alignHorizontal: "right"</code> 和 <code>indentRight: \<value\></code></li></ul><ul><li><code>alignHorizontal: "distributed"</code> 和 <code>indentDistributed: \<value\></code></li></ul>|
| `alignVertical:`|"top" \| "center" \| "bottom" \| "justify" \| "distributed"||



**背景**


|**索引鍵**|**值**|**附註**|
|:-----|:-----|:-----|
| `backgroundColor:`|"none" \| \<所有預先定義的色彩名稱\> \| #RRGGBB|預先定義的色彩名稱：<br/><br/>"black"、"blue"、"gray"、"green"、"orange"、"pink"、"purple"、"red"、"teal"、"turquoise"、"violet"、"white"、"yellow"|



**框線**


|**索引鍵**|**值**|**附註**|
|:-----|:-----|:-----|
| `borderStyle:`|"none" \| \<所有預先定義的框線樣式名稱\>|預先定義的框線樣式名稱：<br/><br/>"dash dot"、"dash dot dot"、"dashed"、"dotted"、"double"、"hair"、"medium dash dot"、"medium dash dot dot"、"medium dashed"、"medium"、"slant dash dot"、"thick"、"thin"<br/><br/>套用至指定範圍內的所有框線。(相當於在 [儲存格格式] 對話方塊的 [框線] 索引標籤上同時使用 [外框] 和 [內線] 預設格式指定框線樣式。)<br/><br/> **附註︰**Excel 2013 支援轉譯所有 13 個預先定義的框線樣式。不過，Excel Online 不支援每個框線樣式。下表說明當您在 Excel Online 中開啟試算表時，用於每一種框線樣式的呈現。<br/><br/><table><tr><th>Excel 2013</th><th>Excel Online</th></tr><tr><td>"dash dot"</td><td>虛線 (1 像素)</td></tr><tr><td>"dash dot dot"</td><td>虛線 (1 像素)</td></tr><tr><td>"dashed"</td><td>虛線 (1 像素)</td></tr><tr><td>"dotted"</td><td>虛線 (1 像素)</td></tr><tr><td>"double"</td><td>雙精確度 (3 像素)</td></tr><tr><td>"hair"</td><td>單色 (1 像素)</td></tr><tr><td>"medium dash dot"</td><td>虛線 (2 像素)</td></tr><tr><td>"medium dash dot dot"</td><td>虛線 (2 像素)</td></tr><tr><td>"medium dashed"</td><td>虛線 (2 像素)</td></tr><tr><td>"medium"</td><td>單色 (2 像素)</td></tr><tr><td>"slant dash dot"</td><td>虛線 (2 像素)</td></tr><tr><td>"thick"</td><td>單色 (3 像素)</td></tr><tr><td>"thin"</td><td>單色 (1 像素)</td></tr></table>|
| `borderColor:`|"automatic" \| \<所有預先定義的色彩名稱\> \| #RRGGBB|套用至指定範圍內的所有框線。|
| `borderTopStyle:`|"none" \| \<所有預先定義的框線樣式名稱\>|套用至指定範圍內的所有框線。|
| `borderTopColor:`|"automatic" \| \<所有預先定義的色彩名稱\> \| #RRGGBB|套用至指定範圍內的所有框線。|
| `borderBottomStyle:`|"none" \| \<所有預先定義的框線樣式名稱\>|套用至指定範圍內的所有框線。|
| `borderBottomColor:`|"automatic" \| \<所有預先定義的色彩名稱\> \| #RRGGBB|套用至指定範圍內的所有框線。|
| `borderLeftStyle:`|"none" \| \<所有預先定義的框線樣式名稱\>|套用至指定範圍內的所有框線。|
| `borderLeftColor:`|"automatic" \| \<所有預先定義的色彩名稱\> \| #RRGGBB|套用至指定範圍內的所有框線。|
| `borderRightStyle:`|"none" \| \<所有預先定義的框線樣式名稱\>|套用至指定範圍內的所有框線。|
| `borderRightColor:`|"automatic" \| \<所有預先定義的色彩名稱\> \| #RRGGBB|套用至指定範圍內的所有框線。|
| `borderOutlineStyle:`|"none" \| \<所有預先定義的框線樣式名稱\>|套用至指定範圍內的所有框線。|
| `borderOutlineColor:`|"automatic" \| \<所有預先定義的色彩名稱\> \| #RRGGBB|套用至指定範圍內的所有框線。|
| `borderInlineStyle:`|"none" \| \<所有預先定義的框線樣式名稱\>|僅套用至指定範圍內的所有框線。(相當於在 [儲存格格式] 對話方塊的 [外框] 索引標籤上只使用 [內線] 預設格式指定框線樣式。)|
| `borderInlineColor:`|"automatic" \| \<所有預先定義的色彩名稱\> \| #RRGGBB|僅套用至指定範圍內的所有框線 |



**儲存格寬度、高度和文繞圖**


|**索引鍵**|**值**|
|:-----|:-----|
| `width:`|"auto fit" \|  **數字**|
| `height:`|"auto fit" \|  **數字**|
| `wrapping:`|**布林值**|



**字型**


|**索引鍵**|**值**|**附註**|
|:-----|:-----|:-----|
| `fontFamily:`|\<所有可用的字型名稱\>|在 Excel Online 中設定字型時，如果字型不在瀏覽器中，API 嘗試回復為下列字型 (依此順序)︰Segoe UI、Thonburi、Arial、Verdana 和 Microsoft Sans Serif 字型。如果這些字型均無法使用，則會使用瀏覽器的預設字型。|
| `fontStyle:`|"regular" \| "italic" \| "bold" \| "bold italic"|**附註**：發行時，將 `fontStyle:` 設定為 "italic"，然後後續設定為 "bold" (反之亦然)，其作用會是這兩項設定的聯集。也就是說，如果您先設定 "italic"，然後接著設定 "bold"，結果會是 "bold italic"。若_只要_在先前已設定為粗體或斜體的某個範圍內設定斜體或粗體中的一個，您必須先設定 `fontStyle:"regular"` 來清除先前的格式設定。|
| `fontSize:`|**數字**||
| `fontUnderlineStyle:`|"none" \| "single" \| "double" \| "single accounting" \| "double accounting"||
| `fontColor:`|"automatic" \| \<所有預先定義的色彩名稱\> \| #RRGGBB||
| `fontDirection:`|"context" \| "left-to-right" \| "right-to-left"|Excel Online 目前不支援從右至左方向顯示文字。不過，如果您的增益集在 Excel Online 中執行時將 `fontDirection:` 設定為「從右至左」，在桌面版 Excel 中開啟活頁簿時，該格式設定會儲存在活頁簿檔案中並正確顯示。|
| `fontStrikethrough:`|**布林值**||
| `fontSuperscript:`|**布林值**||
| `fontSubScript:`|**布林值**||
| `fontNormal:`|**布林值**|將字型、字型樣式、大小和效果設定為一般樣式。這樣會將儲存格字型格式設定重設為預設值。相當於在 [儲存格格式] 對話方塊的 [字型] 索引標籤選取 [一般字型] 核取方塊。|



**縮排**


|**索引鍵**|**值**|**附註**|
|:-----|:-----|:-----|
| `indentLeft:`|**數字**|當與對齊值組合時，僅支援下列的組合︰<br/><br/><ul><li><code>alignHorizontal: "left"</code> 和 <code>indentLeft: \<value\></code></li></ul>|
| `indentRight:`|**數字**|當與對齊值組合時，僅支援下列的組合︰<br/><br/><ul><li><code>alignHorizontal: "right"</code> 和 <code>indentRight: \<value\></code></li></ul>|
| `indentDistributed:`|**數字**|當與對齊值組合時，僅支援下列的組合︰<br/><br/><ul><li><code>alignHorizontal: "distributed"</code> 和 <code>indentDistributed: \<value\></code></li></ul>|



**數字格式**


|**索引鍵**|**值**|**附註**|
|:-----|:-----|:-----|
| `numberFormat:`|**字串**|若要指定數字格式，請使用自訂的數字格式字串。例如，若要指定兩位小數，並以逗號作為千分位分隔符號，您可以指定︰ <br/><br/> `numberFormat:"#,###.00"`<br/><br/>這些是您可以[使用 [儲存格格式] 對話方塊的 [數值] 索引標籤上 [自訂] 格式類別建立](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1)的相同自訂格式字串。<br/><br/> **提示：**您可以利用下列步驟看到 Excel 的 [儲存格格式] 對話方塊中標準目錄的格式字串的外觀︰<br/><br/><ol><li>從 [類別]<b></b> 清單中選取一個標準格式類別，例如 [貨幣]<span class="ui"></span>。</li><li>在對話方塊的右邊設定格式的選項。</li><li>選取 [自訂]<b></b> 類別來檢視 [型別] <b></b> 清單頂端的格式字串。</li></ol>|



**表格選項**


|**索引鍵**|**值**|**附註**|
|:-----|:-----|:-----|
| `style:`|"none" \| \<所有預先定義的表格樣式名稱\>|預先定義的表格樣式名稱：<br/><br/>"TableStyleLight1"、"TableStyleLight2"、"TableStyleLight3"、"TableStyleLight4"、"TableStyleLight5"、"TableStyleLight6"、"TableStyleLight7"、"TableStyleLight8"、"TableStyleLight9"、"TableStyleLight10"、"TableStyleLight11"、"TableStyleLight12"、"TableStyleLight13"、"TableStyleLight14"、"TableStyleLight15"、"TableStyleLight16"、"TableStyleLight17"、"TableStyleLight18"、"TableStyleLight19"、"TableStyleLight20"、"TableStyleLight21"、"TableStyleMedium1"、"TableStyleMedium2"、"TableStyleMedium3"、"TableStyleMedium4"、"TableStyleMedium5"、"TableStyleMedium6"、"TableStyleMedium7"、"TableStyleMedium8"、"TableStyleMedium9"、"TableStyleMedium10"、"TableStyleMedium11"、"TableStyleMedium12"、"TableStyleMedium13"、"TableStyleMedium14"、"TableStyleMedium15"、"TableStyleMedium16"、"TableStyleMedium17"、"TableStyleMedium18"、"TableStyleMedium19"、"TableStyleMedium20"、"TableStyleMedium21"、"TableStyleMedium22"、"TableStyleMedium23"、"TableStyleMedium24"、"TableStyleMedium25"、"TableStyleMedium26"、"TableStyleMedium27"、"TableStyleMedium28"、"TableStyleDark1"、"TableStyleDark2"、"TableStyleDark3"、"TableStyleDark4"、"TableStyleDark5"、"TableStyleDark6"、"TableStyleDark7"、"TableStyleDark8"、"TableStyleDark9"、"TableStyleDark10"、"TableStyleDark11"<br/><br/>若要查看表格樣式的外觀，請在 Excel 中插入表格，在 [表格工具] \| [設計] 索引標籤上，選擇 [快速樣式] 下拉式功能表，然後選取預先定義的樣式。樣式的工具提示會對應到上述清單中的其中一個值。|
| `headerRow:`|**布林值**||
| `firstColumn:`|**布林值**||
| `filterButton:`|**布林值**||
| `totalRow:`|**布林值**||
| `lastColumn:`|**布林值**||
| `bandedRows:`|**布林值**||
| `bandedColumns:`|**布林值**||
