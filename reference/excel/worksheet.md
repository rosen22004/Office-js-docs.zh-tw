# <a name="worksheet-object-(javascript-api-for-excel)"></a>Worksheet 物件 (適用於 Excel 的 JavaScript API)

Excel 工作表是一組儲存格方格。它可以包含資料、表格、圖表等。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|id|string|傳回可在特定活頁簿中唯一識別工作表的值。即使重新命名或移動工作表，識別碼的值仍保持不變。值會隨著開啟之檔案的每個工作階段不同而改變。唯讀。|
|name|string|工作表的顯示名稱。|
|position|int|活頁簿內以零起始的工作表位置。|
|visibility|字串|工作表的可見度。可能的值為：Visible、Hidden、VeryHidden。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|charts|[ChartCollection](chartcollection.md)|代表屬於活頁簿一部份的圖表集合。唯讀。|
|protection|[WorksheetProtection](worksheetprotection.md)|傳回工作表的工作表保護物件。唯讀。|
|tables|[TableCollection](tablecollection.md)|代表屬於活頁簿一部份的表格集合。唯讀。|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|在 Excel UI 中啟動工作表。|
|[delete()](#delete)|void|從活頁簿中刪除工作表。|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|根據列和欄數，取得包含單一儲存格的 range 物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|取得由位址或名稱指定的 range 物件。|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|使用的範圍是最小範圍，其中包含具有值或獲指派格式設定的任何儲存格。如果工作表空白，則此函數會傳回左上角儲存格。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


### <a name="activate()"></a>activate()
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
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="delete()"></a>delete()
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
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcell(row:-number,-column:-number)"></a>getCell(row: number, column: number)
根據列和欄數，取得包含單一儲存格的 range 物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。

#### <a name="syntax"></a>語法
```js
worksheetObject.getCell(row, column);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
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
    var cell = worksheet.getCell(0,0);
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


### <a name="getrange(address:-string)"></a>getRange(address: string)
取得由位址或名稱指定的 range 物件。

#### <a name="syntax"></a>語法
```js
worksheetObject.getRange(address);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|地址|string|選用。範圍的位址或名稱。如果未指定，則會傳回整個工作表範圍。|

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
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

### <a name="getusedrange(valuesonly:-bool)"></a>getUsedRange(valuesOnly: bool)
使用的範圍是最小範圍，其中包含具有值或獲指派格式設定的任何儲存格。如果工作表空白，則此函數會傳回左上角儲存格。

#### <a name="syntax"></a>語法
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|valuesOnly|bool|選用。僅將包含值的儲存格考慮為使用的儲存格 (忽略格式設定)。|

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


### <a name="load(param:-object)"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
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

