# Worksheet 物件 (適用於 Excel 的 JavaScript API)

Excel 工作表是一組儲存格方格。它可以包含資料、表格、圖表等。

## 屬性

| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|id|string|傳回可在特定活頁簿中唯一識別工作表的值。 即使重新命名或移動工作表，識別碼的值仍保持不變。 值會隨著開啟之檔案的每個工作階段不同而改變。 唯讀。|
|name|string|工作表的顯示名稱。|
|position|int|活頁簿內以零起始的工作表位置。|
|visibility|string|工作表的可見度。可能的值為：Visible、Hidden、VeryHidden。|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
| 關聯性 | 類型	   |說明|
|:---------------|:--------|:----------|
|charts|[ChartCollection](chartcollection.md)|代表屬於活頁簿一部份的圖表集合。唯讀。|
|保護|[WorksheetProtection](worksheetprotection.md)|傳回工作表的工作表保護物件。唯讀。|
|tables|[TableCollection](tablecollection.md)|代表屬於活頁簿一部份的表格集合。唯讀。|

## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|在 Excel UI 中啟動工作表。|
|[delete()](#delete)|void|從活頁簿中刪除工作表。|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[範圍](range.md)|根據列和欄數，取得包含單一儲存格的 range 物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。|
|[getRange(address: string)](#getrangeaddress-string)|[範圍](range.md)|取得由位址或名稱指定的 range 物件。|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[範圍](range.md)|使用的範圍是最小範圍，其中包含具有值或獲指派格式設定的任何儲存格。如果工作表空白，則此函數會傳回左上角儲存格。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料


### activate()
在 Excel UI 中啟動工作表。

#### 語法
```js
worksheetObject.activate();
```

#### 參數
無

#### 傳回
void

#### 範例

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


### delete()
從活頁簿中刪除工作表。

#### 語法
```js
worksheetObject.delete();
```

#### 參數
無

#### 傳回
void

#### 範例

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


### getCell(row: number, column: number)
根據列和欄數，取得包含單一儲存格的 range 物件。只要儲存格保持在工作表方格中，此儲存格可以位於其父範圍的界限之外。

#### 語法
```js
worksheetObject.getCell(row, column);
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


### getRange(address: string)
取得由位址或名稱指定的 range 物件。

#### 語法
```js
worksheetObject.getRange(address);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|地址|string|選用。範圍的位址或名稱。如果未指定，則會傳回整個工作表範圍。|

#### 傳回
[範圍](range.md)

#### 範例
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

### getUsedRange(valuesOnly: bool)
使用的範圍是最小範圍，其中包含具有值或獲指派格式設定的任何儲存格。如果工作表空白，則此函數會傳回左上角儲存格。

#### 語法
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|valuesOnly|bool|選用。僅將包含值的儲存格考慮為使用的儲存格 (忽略格式設定)。|

#### 傳回
[範圍](range.md)

#### 範例

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
### 屬性存取範例

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

