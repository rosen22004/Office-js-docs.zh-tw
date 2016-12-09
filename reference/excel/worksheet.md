# <a name="worksheet-object-javascript-api-for-excel"></a>Worksheet 物件 (適用於 Excel 的 JavaScript API)

Excel 工作表是一組儲存格方格。它可以包含資料、表格、圖表等。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|id|string|傳回可在特定活頁簿中唯一識別工作表的值。即使重新命名或移動工作表，識別碼的值仍保持不變。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|工作表的顯示名稱。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|活頁簿內以零起始的工作表位置。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visibility|string|工作表的可見度。可能的值為：Visible、Hidden、VeryHidden。|[1.1，1.1 版可閱讀可見度；1.2 版則可進行設定。](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|charts|[ChartCollection](chartcollection.md)|代表屬於活頁簿一部份的圖表集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|表示屬於活頁簿一部份的樞紐分析表集合。唯讀。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[WorksheetProtection](worksheetprotection.md)|傳回工作表的工作表保護物件。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|代表屬於活頁簿一部份的表格集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[activate()](#activate)|void|在 Excel UI 中啟動工作表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|從活頁簿中刪除工作表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|根據列和欄數，取得包含單一儲存格的範圍物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|取得由位址或名稱指定的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|使用的範圍是最小範圍，其中包含具有值或獲指派格式設定的任何儲存格。如果工作表空白，則此函數會傳回左上角儲存格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="activate"></a>activate()
在 Excel UI 中啟動工作表。

#### <a name="syntax"></a>語法
```js
worksheetObject.activate();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="delete"></a>delete()
從活頁簿中刪除工作表。

#### <a name="syntax"></a>語法
```js
worksheetObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
根據列和欄數，取得包含單一儲存格的範圍物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。

#### <a name="syntax"></a>語法
```js
worksheetObject.getCell(row, column);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|列|number|要擷取之儲存格的列號。以 0 開始編製索引。|
|column|number|要擷取之儲存格的欄數。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrangeaddress-string"></a>getRange(address: string)
取得由位址或名稱指定的範圍物件。

#### <a name="syntax"></a>語法
```js
worksheetObject.getRange(address);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|地址|string|選用。範圍的位址或名稱。如果未指定，則會傳回整個工作表範圍。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
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
    var sheetName = "Sheet1";
    var rangeName = 'MyRange';
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getusedrangevaluesonly"></a>getUsedRange(valuesOnly)
使用的範圍是最小範圍，其中包含具有值或獲指派格式設定的任何儲存格。如果工作表空白，則此函數會傳回左上角儲存格。

#### <a name="syntax"></a>語法
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|僅將包含值的儲存格考慮為使用的儲存格 (忽略格式設定)。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    var usedRange = worksheet.getUsedRange();
    usedRange.load('address');
    return ctx.sync().then(function() {
            console.log(usedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="loadparam-object"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例

根據工作表名稱取得工作表屬性。

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.load('position')
    return ctx.sync().then(function() {
            console.log(worksheet.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

設定工作表位置。 

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.position = 2;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
