# Range 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Excel for iOS、Office 2016_

Range 代表一組一或多個連續儲存格，例如儲存格、列、欄或儲存格區塊等。

## 屬性

| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|地址|string|代表 A1 樣式的範圍參照。位址值會包含工作表參照 (例如 Sheet1!A1: B4)。唯讀。|
|addressLocal|string|以使用者的語言表示指定範圍的範圍參照。唯讀。|
|cellCount|int|範圍中的儲存格數目。唯讀。|
|columnCount|int|代表範圍中的欄總數。唯讀。|
|columnHidden|bool|表示是否隱藏目前範圍的所有資料行。|
|columnIndex|int|代表範圍中第一個儲存格的欄號。以 0 開始編製索引。唯讀。|
|formulas|object[]|代表 A1 樣式標記法的公式。|
|formulasLocal|object[][]|以使用者的語言和數字格式地區設定，表示 A1 樣式標記法的公式。例如，英文的 "=SUM(A1, 1.5)" 公式在德文中會表示為 "=SUMME(A1; 1,5)"。|
|formulasR1C1|object[][]|代表 R1C1 樣式標記法的公式。|
|hidden|bool|表示是否隱藏目前範圍的所有儲存格。唯讀。|
|numberFormat|object[][]|代表特定儲存格的數字格式代碼。|
|rowCount|int|傳回範圍中的列總數。唯讀。|
|rowHidden|bool|表示是否隱藏目前範圍的所有資料列。|
|rowIndex|int|傳回範圍中第一個儲存格的列號。以 0 開始編製索引。唯讀。|
|文字|object[][]|所指定範圍的文字值。文字值與儲存格寬度無關。Excel UI 中出現的 # 替代符號不會影響 API 所傳回的文字值。唯讀。|
|valueTypes|string|代表每個儲存格的資料類型。唯讀。可能的值為：Unknown、Empty、String、Integer、Double、Boolean、Error。|
|values|object[][]|代表所指定範圍的原始值。傳回的資料可能是 string、number 或 boolean 類型。包含錯誤的儲存格會傳回錯誤字串。|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
| 關聯性 | 類型	   |說明|
|:---------------|:--------|:----------|
|format|[RangeFormat](rangeformat.md)|傳回格式物件，其中封裝了範圍的字型、填滿、框線、對齊方式及其他屬性。唯讀。|
|排序|[RangeSort](rangesort.md)|代表範圍的排序組態。唯讀。|
|工作表|[Worksheet](worksheet.md)|包含目前範圍的工作表。唯讀。|

## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[clear(applyTo: string)](#clearapplyto-string)|void|清除範圍值、格式、填滿、框線等。|
|[delete(shift: string)](#deleteshift-string)|void|刪除範圍相關的儲存格。|
|[getBoundingRect(anotherRange:Range or string)](#getboundingrectanotherrange-range-or-string)|[範圍](range.md)|取得包含特定範圍的最小 range 物件。例如，"B2:C5" 和 "D10:E15" 的 getBoundingRect 是 "B2:E15"。|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[範圍](range.md)|根據列和欄數，取得包含單一儲存格的 range 物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。傳回的儲存格位置相對於範圍的左上角儲存格。|
|[getColumn(column: number)](#getcolumncolumn-number)|[範圍](range.md)|取得範圍內包含的欄。|
|[getEntireColumn()](#getentirecolumn)|[範圍](range.md)|取得物件，代表範圍的整個欄。|
|[getEntireRow()](#getentirerow)|[範圍](range.md)|取得物件，代表範圍的整個列。|
|[getIntersection(anotherRange:Range or string)](#getintersectionanotherrange-range-or-string)|[範圍](range.md)|取得 range 物件，代表特定範圍的矩形交集。|
|[getLastCell()](#getlastcell)|[範圍](range.md)|取得範圍內最後一個儲存格。例如，"B2:D5" 的最後一個儲存格是 "D5"。|
|[getLastColumn()](#getlastcolumn)|[範圍](range.md)|取得範圍內最後一欄。例如，"B2:D5" 的最後一欄是 "D2:D5"。|
|[getLastRow()](#getlastrow)|[範圍](range.md)|取得範圍內最後一列。例如，"B2:D5" 的最後一列是 "B5:D5"。|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[範圍](range.md)|取得物件，代表從指定範圍偏移的範圍。傳回範圍的維度會符合此範圍。如果產生的範圍強制超出工作表方格的界限，則會擲回例外狀況。|
|[getRow(row: number)](#getrowrow-number)|[範圍](range.md)|取得範圍內包含的列。|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[範圍](range.md)|傳回 range 物件的使用的子範圍。|
|[insert(shift: string)](#insertshift-string)|[範圍](range.md)|在工作表中插入一個儲存格或儲存格範圍以取代此範圍，並移動其他儲存格以挪出空間。傳回位於現在空格的新 Range 物件。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|
|[merge(across: bool)](#mergeacross-bool)|void|合併範圍儲存格到工作表中的一個區域。|
|[select()](#select)|void|在 Excel UI 中選取指定的範圍。|
|[unmerge()](#unmerge)|void|取消將範圍儲存格合併至個別儲存格。|

## 方法詳細資料


### clear(applyTo: string)
清除範圍值、格式、填滿、框線等。

#### 語法
```js
rangeObject.clear(applyTo);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|applyTo|string|選用。決定清除動作的類型。可能的值為：`All` 預設選項、`Formats`、`Contents`。|

#### 傳回
void

#### 範例

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


### delete(shift: string)
刪除範圍相關的儲存格。

#### 語法
```js
rangeObject.delete(shift);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|SHIFT|string|指定移動儲存格的方式。可能的值為：Up、Left。|

#### 傳回
void

#### 範例

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


### getBoundingRect(anotherRange:Range or string)
取得包含特定範圍的最小 range 物件。例如，"B2:C5" 和 "D10:E15" 的 GetBoundingRect 是 "B2:E15"。

#### 語法
```js
rangeObject.getBoundingRect(anotherRange);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|anotherRange|Range 或 string|Range 物件或位址或範圍名稱。|

#### 傳回
[範圍](range.md)

#### 範例

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


### getCell(row: number, column: number)
根據列和欄數，取得包含單一儲存格的 range 物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。傳回的儲存格位置相對於範圍的左上角儲存格。

#### 語法
```js
rangeObject.getCell(row, column);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|列|number|要擷取之儲存格的列號。以 0 開始編製索引。|
|column|number|要擷取之儲存格的欄號。以 0 開始編製索引。|

#### 傳回
[範圍](range.md)

#### 範例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var cell = range.getCell(0,0);
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


### getColumn(column: number)
取得範圍內包含的欄。

#### 語法
```js
rangeObject.getColumn(column);
```

#### 參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|column|number|要擷取之範圍的欄號。以 0 開始編製索引。|

#### 傳回
[範圍](range.md)

#### 範例

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


### getEntireColumn()
取得物件，代表範圍的整個欄。

#### 語法
```js
rangeObject.getEntireColumn();
```

#### 參數
無

#### 傳回
[範圍](range.md)

#### 範例

附註：Range 的方格屬性 (values、numberFormat、formulas) 包含 `null`，因為相關範圍為無界限。

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

### getEntireRow()
取得物件，代表範圍的整個列。

#### 語法
```js
rangeObject.getEntireRow();
```

#### 參數
無

#### 傳回
[範圍](range.md)

#### 範例
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
Range 的方格屬性 (values、numberFormat、formulas) 包含 `null`，因為相關範圍為無界限。

### getIntersection(anotherRange:Range or string)
取得 range 物件，代表特定範圍的矩形交集。

#### 語法
```js
rangeObject.getIntersection(anotherRange);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|anotherRange|Range 或 string|將用來決定範圍交集的 Range 物件或範圍位址。|

#### 傳回
[範圍](range.md)

#### 範例

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


### getLastCell()
取得範圍內最後一個儲存格。例如，"B2:D5" 的最後一個儲存格是 "D5"。

#### 語法
```js
rangeObject.getLastCell();
```

#### 參數
無

#### 傳回
[範圍](range.md)

#### 範例

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


### getLastColumn()
取得範圍內最後一欄。例如，"B2:D5" 的最後一欄是 "D2:D5"。

#### 語法
```js
rangeObject.getLastColumn();
```

#### 參數
無

#### 傳回
[範圍](range.md)

#### 範例

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


### getLastRow()
取得範圍內最後一列。例如，"B2:D5" 的最後一列是 "B5:D5"。

#### 語法
```js
rangeObject.getLastRow();
```

#### 參數
無

#### 傳回
[範圍](range.md)

#### 範例

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



### getOffsetRange(rowOffset: number, columnOffset: number)
取得物件，代表從指定範圍偏移的範圍。傳回範圍的維度會符合此範圍。如果產生的範圍強制超出工作表方格的界限，則會擲回例外狀況。

#### 語法
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|rowOffset|number|該範圍要偏移的列數 (正值、負值或 0)。正值表示向下偏移，負值表示向上偏移。|
|columnOffset|number|該範圍要偏移的欄數 (正值、負值或 0)。正值表示向右偏移，負值表示向左偏移。|

#### 傳回
[範圍](range.md)

#### 範例

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


### getRow(row: number)
取得範圍內包含的列。

#### 語法
```js
rangeObject.getRow(row);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|列|number|要擷取之範圍的列號。以 0 開始編製索引。|

#### 傳回
[範圍](range.md)

#### 範例

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


### getUsedRange(valuesOnly: bool)
傳回特定 range 物件所使用的範圍。

#### 語法
```js
rangeObject.getUsedRange(valuesOnly);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|valuesOnly|bool|選用。若為 true，只會將目前有值的儲存格視為使用的儲存格。預設為 false，會將任何曾具有值的儲存格視為已使用。|

#### 傳回
[範圍](range.md)

#### 範例

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


### insert(shift: string)
在工作表中插入一個儲存格或儲存格範圍以取代此範圍，並移動其他儲存格以挪出空間。傳回位於現在空格的新 Range 物件。

#### 語法
```js
rangeObject.insert(shift);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|SHIFT|string|指定移動儲存格的方式。可能的值為：Down、Right。|

#### 傳回
[範圍](range.md)

#### 範例

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


### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void

### merge(across: bool)
合併範圍儲存格到工作表中的一個區域。

#### 語法
```js
rangeObject.merge(across);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|跨列|bool|選用。若設為 True，則會將指定範圍的每一列的儲存格合併成個別的合併儲存格。預設值為 False。|

#### 傳回
void

#### 範例
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


### select()
在 Excel UI 中選取指定的範圍。

#### 語法
```js
rangeObject.select();
```

#### 參數
無

#### 傳回
void

#### 範例

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### unmerge()
取消將合併的儲存格範圍合併至個別儲存格。

#### 語法
```js
rangeObject.unmerge();
```

#### 參數
無

#### 傳回
void

#### 範例
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

### 屬性存取範例

這個範例會使用範圍位址，以取得 range 物件。

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

這個範例會使用具名範圍，以取得 range 物件。

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
下列範例與上方的範例相同，不同之處在於它使用公式的 R1C1 表示法。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulasR1C1 = [[null,null], [null,null], [null,"=R[-1]C-R[-2]C"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulasR1C1= formulasR1C1;
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
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    range = namedItem.range;
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

