# TableColumn 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表表格中的一欄。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|id|int|傳回可在表格中識別欄的唯一索引鍵。唯讀。|
|index|int|傳回表格欄集合中欄的索引編號。以 0 開始編製索引。唯讀。|
|name|string|傳回表格欄的名稱。唯讀。|
|values|object[][]|代表所指定範圍的原始值。傳回的資料可能是 string、number 或 boolean 類型。包含錯誤的儲存格會傳回錯誤字串。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|從表格中刪除欄。|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|取得與欄的資料主體相關的 range 物件。|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|取得與欄的標題列相關的 range 物件。|
|[getRange()](#getrange)|[Range](range.md)|取得與整個欄相關的 range 物件。|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|取得與欄的合計列相關的 range 物件。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### delete()
從表格中刪除欄。

#### 語法
```js
tableColumnObject.delete();
```

#### 參數
無

#### 傳回
void

#### 範例

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
	column.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getDataBodyRange()
取得與欄的資料主體相關的 range 物件。

#### 語法
```js
tableColumnObject.getDataBodyRange();
```

#### 參數
無

#### 會傳回
[Range](range.md)

#### 範例

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var dataBodyRange = column.getDataBodyRange();
	dataBodyRange.load('address');
	return ctx.sync().then(function() {
		console.log(dataBodyRange.address);
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getHeaderRowRange()
取得與欄的標題列相關的 range 物件。

#### 語法
```js
tableColumnObject.getHeaderRowRange();
```

#### 參數
無

#### 會傳回
[Range](range.md)

#### 範例

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var headerRowRange = columns.getHeaderRowRange();
	headerRowRange.load('address');
	return ctx.sync().then(function() {
		console.log(headerRowRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getRange()
取得與整個欄相關的 range 物件。

#### 語法
```js
tableColumnObject.getRange();
```

#### 參數
無

#### 會傳回
[Range](range.md)

#### 範例

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var columnRange = columns.getRange();
	columnRange.load('address');
	return ctx.sync().then(function() {
		console.log(columnRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getTotalRowRange()
取得與欄的合計列相關的 range 物件。

#### 語法
```js
tableColumnObject.getTotalRowRange();
```

#### 參數
無

#### 會傳回
[Range](range.md)

#### 範例

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var totalRowRange = columns.getTotalRowRange();
	totalRowRange.load('address');
	return ctx.sync().then(function() {
		console.log(totalRowRange.address);
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
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|param|object|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void
### 屬性存取範例

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItem(0);
	column.load('index');
	return ctx.sync().then(function() {
		console.log(column.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
	column.values = newValues;
	column.load('values');
	return ctx.sync().then(function() {
		console.log(column.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
