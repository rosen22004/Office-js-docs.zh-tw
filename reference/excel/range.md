# <a name="range-object-javascript-api-for-excel"></a>Range 物件 (適用於 Excel 的 JavaScript API)

Range 代表一組一或多個連續儲存格，例如儲存格、列、欄或儲存格區塊等。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|地址|string|代表 A1 樣式的範圍參照。位址值會包含工作表參照 (例如 Sheet1!A1: B4)。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|addressLocal|string|以使用者的語言表示指定範圍的範圍參照。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|cellCount|Int|範圍中的儲存格數目。如果儲存格計數超過 2^31-1 (2,147,483,647)，此 API 將會傳回 -1。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|代表範圍中的欄總數。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnHidden|bool|表示是否隱藏目前範圍的所有資料行。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|columnIndex|int|代表範圍中第一個儲存格的欄號。以 0 開始編製索引。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|代表 A1 樣式標記法的公式。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|以使用者的語言和數字格式地區設定，表示 A1 樣式標記法的公式。例如，英文的 "=SUM(A1, 1.5)" 公式在德文中會表示為 "=SUMME(A1; 1,5)"。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|代表 R1C1 樣式標記法的公式。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|hidden|bool|表示是否隱藏目前範圍的所有儲存格。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|代表特定儲存格的 Excel 數字格式代碼。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|傳回範圍中的列總數。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHidden|bool|表示是否隱藏目前範圍的所有資料列。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|rowIndex|int|傳回範圍中第一個儲存格的列號。以 0 開始編製索引。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|object[][]|所指定範圍的文字值。文字值與儲存格寬度無關。Excel UI 中出現的 # 替代符號不會影響 API 所傳回的文字值。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|代表每個儲存格的資料類型。唯讀。可能的值為：Unknown、Empty、String、Integer、Double、Boolean、Error。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|代表所指定範圍的原始值。傳回的資料可能是 string、number 或 boolean 類型。包含錯誤的儲存格會傳回錯誤字串。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|format|[RangeFormat](rangeformat.md)|傳回格式物件，其中封裝了範圍的字型、填滿、框線、對齊方式及其他屬性。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[RangeSort](rangesort.md)|代表目前範圍的範圍排序。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|包含目前範圍的工作表。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[clear(applyTo: string)](#clearapplyto-string)|void|清除範圍值、格式、填滿、框線等。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete(shift: string)](#deleteshift-string)|void|刪除範圍相關的儲存格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getBoundingRect(anotherRange:Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|取得包含特定範圍的最小範圍物件。例如，"B2:C5" 和 "D10:E15" 的 GetBoundingRect 是 "B2:E16"。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|根據列和欄數，取得包含單一儲存格的範圍物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。傳回的儲存格位置相對於範圍的左上角儲存格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|取得範圍內包含的欄。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsAfter(count: number)](#getcolumnsaftercount-number)|[Range](range.md)|取得目前的 Range 物件右邊的欄數。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsBefore(count: number)](#getcolumnsbeforecount-number)|[Range](range.md)|取得目前的 Range 物件左邊的欄數。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireColumn()](#getentirecolumn)|[範圍](range.md)|取得代表範圍整個資料欄的物件 (例如，如果目前範圍代表儲存格 "B4:E11" 時，它的 `getEntireColumn` 則是代表資料欄 "B:E" 的範圍)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRow()](#getentirerow)|[範圍](range.md)|取得代表範圍整個資料行的物件 (例如，如果目前範圍代表儲存格 "B4:E11" 時，它的 `GetEntireRow` 則是代表資料列 "4:11" 的範圍)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersection(anotherRange:Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|取得範圍物件，代表特定範圍的矩形交集。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersectionOrNullObject(anotherRange:Range 或 string)](#getintersectionornullobjectanotherrange-range-or-string)|[Range](range.md)|取得範圍物件，代表特定範圍的矩形交集。如果找到沒有交集，則會傳回 null 物件。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastCell()](#getlastcell)|[Range](range.md)|取得範圍內最後一個儲存格。例如，"B2:D5" 的最後一個儲存格是 "D5"。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|取得範圍內最後一欄。例如，"B2:D5" 的最後一欄是 "D2:D5"。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastRow()](#getlastrow)|[Range](range.md)|取得範圍內最後一列。例如，"B2:D5" 的最後一列是 "B5:D5"。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|取得物件，代表從指定範圍偏移的範圍。傳回範圍的維度會符合此範圍。如果產生的範圍強制超出工作表方格的界限，則將會擲回錯誤。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getResizedRange(deltaRows: number, deltaColumns: number)](#getresizedrangedeltarows-number-deltacolumns-number)|[Range](range.md)|取得與目前 Range 物件類似的 Range 物件，但右下角以一定的欄與列數展開 (或收起)。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|取得範圍內包含的列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsAbove(count: number)](#getrowsabovecount-number)|[Range](range.md)|取得目前的 Range 物件上方的列數。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsBelow(count: number)](#getrowsbelowcount-number)|[Range](range.md)|取得目前的 Range 物件下方的列數。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: [ApiSet(Version)](#getusedrangevaluesonly-apisetversion)|[範圍](range.md)|傳回特定 Range 物件所使用的範圍。如果範圍內沒有已使用的儲存格，則此函數會擲回 ItemNotFound 錯誤。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[範圍](range.md)|傳回特定 Range 物件所使用的範圍。如果範圍內沒有已使用的儲存格，則此函數會傳回 null 物件。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getVisibleView()](#getvisibleview)|[RangeView](rangeview.md)|代表目前範圍的可見資料列。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|在工作表中插入一個儲存格或儲存格範圍以取代此範圍，並移動其他儲存格以挪出空間。傳回位於現在空格的新 Range 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[merge(across: bool)](#mergeacross-bool)|void|合併範圍儲存格到工作表中的一個區域。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[select()](#select)|void|在 Excel UI 中選取指定的範圍。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[unmerge()](#unmerge)|void|取消將範圍儲存格合併至個別儲存格。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="clearapplyto-string"></a>clear(applyTo: string)
清除範圍值、格式、填滿、框線等。

#### <a name="syntax"></a>語法
```js
rangeObject.clear(applyTo);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|applyTo|string|選用。決定清除動作的類型。可能的值為：`All` 預設選項、`Formats`、`Contents` |

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

下列範例會清除範圍的格式和內容。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="deleteshift-string"></a>delete(shift: string)
刪除範圍相關的儲存格。

#### <a name="syntax"></a>語法
```js
rangeObject.delete(shift);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|SHIFT|string|指定移動儲存格的方式。可能的值為：Up、Left|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getboundingrectanotherrange-range-or-string"></a>getBoundingRect(anotherRange:Range or string)
取得包含特定範圍的最小範圍物件。例如，"B2:C5" 和 "D10:E15" 的 GetBoundingRect 是 "B2:E16"。

#### <a name="syntax"></a>語法
```js
rangeObject.getBoundingRect(anotherRange);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|anotherRange|Range 或 string|Range 物件或位址或範圍名稱。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var range = range.getBoundingRect("G4:H8");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // Prints Sheet1!D4:H8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
根據列和欄數，取得包含單一儲存格的範圍物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。傳回的儲存格位置相對於範圍的左上角儲存格。

#### <a name="syntax"></a>語法
```js
rangeObject.getCell(row, column);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|列|number|要擷取之儲存格的列號。以 0 開始編製索引。|
|column|number|要擷取之儲存格的欄號。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var cell = range.cell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcolumncolumn-number"></a>getColumn(column: number)
取得範圍內包含的欄。

#### <a name="syntax"></a>語法
```js
rangeObject.getColumn(column);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|column|number|要擷取之範圍的欄號。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet19";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!B1:B8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcolumnsaftercount-number"></a>getColumnsAfter(count: number)
取得目前的 Range 物件右邊的欄數。

#### <a name="syntax"></a>語法
```js
rangeObject.getColumnsAfter(count);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Count|數字|選用。要包含在結果範圍中的欄數。一般情況下，請使用正數建立目前範圍以外的範圍。您也可以使用負數建立目前範圍內的範圍。預設值為 1。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="getcolumnsbeforecount-number"></a>getColumnsBefore(count: number)
取得目前的 Range 物件左邊的欄數。

#### <a name="syntax"></a>語法
```js
rangeObject.getColumnsBefore(count);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Count|數字|選用。要包含在結果範圍中的欄數。一般情況下，請使用正數建立目前範圍以外的範圍。您也可以使用負數建立目前範圍內的範圍。預設值為 1。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="getentirecolumn"></a>getEntireColumn()
取得代表範圍整個資料欄的物件 (例如，如果目前範圍代表儲存格 "B4:E11" 時，它的 `getEntireColumn` 則是代表資料欄 "B:E" 的範圍)。

#### <a name="syntax"></a>語法
```js
rangeObject.getEntireColumn();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

附註：因為相關範圍為無界限，所以 Range 的方格屬性 (values、numberFormat、formulas) 會包含 `null`。

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeEC = range.getEntireColumn();
    rangeEC.load('address');
    return ctx.sync().then(function() {
        console.log(rangeEC.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getentirerow"></a>getEntireRow()
取得代表範圍整個資料行的物件 (例如，如果目前範圍代表儲存格 "B4:E11" 時，它的 `GetEntireRow` 則是代表資料列 "4:11" 的範圍)。

#### <a name="syntax"></a>語法
```js
rangeObject.getEntireRow();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "D:F"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeER = range.getEntireRow();
    rangeER.load('address');
    return ctx.sync().then(function() {
        console.log(rangeER.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
因為相關範圍為無界限，所以 Range 的方格屬性 (values、numberFormat、formulas) 會包含 `null`。


### <a name="getintersectionanotherrange-range-or-string"></a>getIntersection(anotherRange:Range or string)
取得 range 物件，代表特定範圍的矩形交集。

#### <a name="syntax"></a>語法
```js
rangeObject.getIntersection(anotherRange);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|anotherRange|Range 或 string|將用來決定範圍交集的 Range 物件或範圍位址。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!D4:F6
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getintersectionornullobjectanotherrange-range-or-string"></a>getIntersectionOrNullObject(anotherRange:Range 或 string)
取得範圍物件，代表特定範圍的矩形交集。如果找到沒有交集，則會傳回 null 物件。

#### <a name="syntax"></a>語法
```js
rangeObject.getIntersectionOrNullObject(anotherRange);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|anotherRange|Range 或 string|將用來決定範圍交集的 Range 物件或範圍位址。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="getlastcell"></a>getLastCell()
取得範圍內最後一個儲存格。例如，"B2:D5" 的最後一個儲存格是 "D5"。

#### <a name="syntax"></a>語法
```js
rangeObject.getLastCell();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getlastcolumn"></a>getLastColumn()
取得範圍內最後一欄。例如，"B2:D5" 的最後一欄是 "D2:D5"。

#### <a name="syntax"></a>語法
```js
rangeObject.getLastColumn();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F1:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getlastrow"></a>getLastRow()
取得範圍內最後一列。例如，"B2:D5" 的最後一列是 "B5:D5"。

#### <a name="syntax"></a>語法
```js
rangeObject.getLastRow();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A8:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



### <a name="getoffsetrangerowoffset-number-columnoffset-number"></a>getOffsetRange(rowOffset: number, columnOffset: number)
取得物件，代表從指定範圍偏移的範圍。傳回範圍的維度會符合此範圍。如果產生的範圍強制超出工作表方格的界限，則將會擲回錯誤。

#### <a name="syntax"></a>語法
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|rowOffset|number|該範圍要偏移的列數 (正值、負值或 0)。正值表示向下偏移，負值表示向上偏移。|
|columnOffset|number|該範圍要偏移的欄數 (正值、負值或 0)。正值表示向右偏移，負值表示向左偏移。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:F6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!H3:K5
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getresizedrangedeltarows-number-deltacolumns-number"></a>getResizedRange(deltaRows: number, deltaColumns: number)
取得與目前 Range 物件類似的 Range 物件，但右下角以一定的欄與列數展開 (或收起)。

#### <a name="syntax"></a>語法
```js
rangeObject.getResizedRange(deltaRows, deltaColumns);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|deltaRows|數字|將在右下角展開的列數，相對於目前範圍。使用正數來展開範圍，或負數來減少範圍。|
|deltaColumns|數字|將在右下角展開的欄數，相對於目前範圍。使用正數來展開範圍，或負數來減少範圍。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="getrowrow-number"></a>getRow(row: number)
取得範圍內包含的列。

#### <a name="syntax"></a>語法
```js
rangeObject.getRow(row);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|列|number|要擷取之範圍的列號。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A2:F2
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrowsabovecount-number"></a>getRowsAbove(count: number)
取得目前的 Range 物件上方的列數。

#### <a name="syntax"></a>語法
```js
rangeObject.getRowsAbove(count);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Count|數字|選用。要包含在結果範圍中的列數。一般情況下，請使用正數建立目前範圍以外的範圍。您也可以使用負數建立目前範圍內的範圍。預設值為 1。|

#### <a name="returns"></a>傳回
[Range](range.md)

### <a name="getrowsbelowcount-number"></a>getRowsBelow(count: number)
取得目前的 Range 物件下方的列數。

#### <a name="syntax"></a>語法
```js
rangeObject.getRowsBelow(count);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Count|數字|選用。要包含在結果範圍中的列數。一般情況下，請使用正數建立目前範圍以外的範圍。您也可以使用負數建立目前範圍內的範圍。預設值為 1。|

#### <a name="returns"></a>傳回
[範圍](range.md)

### <a name="getusedrangevaluesonly-apisetversion"></a>getUsedRange(valuesOnly: [ApiSet(Version)
傳回特定 Range 物件所使用的範圍。如果範圍內沒有已使用的儲存格，則此函數會擲回 ItemNotFound 錯誤。

#### <a name="syntax"></a>語法
```js
rangeObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|僅將包含值的儲存格考慮為使用的儲存格。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeUR = range.getUsedRange();
    rangeUR.load('address');
    return ctx.sync().then(function() {
        console.log(rangeUR.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getusedrangeornullobjectvaluesonly-bool"></a>getUsedRangeOrNullObject(valuesOnly: bool)
傳回特定 Range 物件所使用的範圍。如果範圍內沒有已使用的儲存格，則此函數會傳回 null 物件。

#### <a name="syntax"></a>語法
```js
rangeObject.getUsedRangeOrNullObject(valuesOnly);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|選用。僅將包含值的儲存格考慮為使用的儲存格。|

#### <a name="returns"></a>傳回
[範圍](range.md)

### <a name="getvisibleview"></a>getVisibleView()
代表目前範圍的可見資料列。

#### <a name="syntax"></a>語法
```js
rangeObject.getVisibleView();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[RangeView](rangeview.md)

### <a name="insertshift-string"></a>insert(shift: string)
在工作表中插入一個儲存格或儲存格範圍以取代此範圍，並移動其他儲存格以挪出空間。傳回位於現在空格的新 Range 物件。

#### <a name="syntax"></a>語法
```js
rangeObject.insert(shift);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|SHIFT|string|指定移動儲存格的方式。可能的值為：Down、Right|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
    
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.insert();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="mergeacross-bool"></a>merge(across: bool)
合併範圍儲存格到工作表中的一個區域。

#### <a name="syntax"></a>語法
```js
rangeObject.merge(across);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|跨列|bool|選用。若設為 True，則會將指定範圍的每一列的儲存格合併成個別的合併儲存格。預設值為 False。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.merge(true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="select"></a>select()
在 Excel UI 中選取指定的範圍。

#### <a name="syntax"></a>語法
```js
rangeObject.select();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="unmerge"></a>unmerge()
取消將範圍儲存格合併至個別儲存格。

#### <a name="syntax"></a>語法
```js
rangeObject.unmerge();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>屬性存取範例

下列範例會使用範圍位址，以取得範圍物件。

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8"; 
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

下列範例會使用具名範圍，以取得範圍物件。

```js

Excel.run(function (ctx) { 
    var rangeName = 'MyRange';
    var range = ctx.workbook.names.getItem(rangeName).range;
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

下列範例會在包含 2x3 方格的方格上設定 numberFormat、values 和 formulas。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulas= formulas;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
取得包含目前範圍的工作表。 

```js
/* This might be broken still - it was broken before because it 
    it was missing 'var', but might still be wrong because of
    getting information without loading properly. */
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    var range = namedItem.range;
    var rangeWorksheet = range.worksheet;
    rangeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(rangeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

