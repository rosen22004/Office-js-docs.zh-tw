# WorksheetCollection 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表屬於活頁簿一部份的 worksheet 物件集合。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|項目|[Worksheet[]](worksheet.md)|Worksheet 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|將新的工作表加入活頁簿中。工作表會加入至現有工作表的結尾處。如果您想要啟動新加入的工作表，請對其呼叫 activate()。|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|取得活頁簿中目前作用中的工作表。|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|使用其名稱或 ID 取得 worksheet 物件。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### add(name: string)
將新的工作表加入活頁簿中。工作表會加入至現有工作表的結尾處。如果您想要啟動新加入的工作表，請對其呼叫 activate()。

#### 語法
```js
worksheetCollectionObject.add(name);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|Name|string|選用。要加入的工作表的名稱。若指定，名稱應為唯一的。若不指定，Excel 會自行決定新工作表的名稱。|

#### 傳回
[Worksheet](worksheet.md)

#### 範例

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sample Name';
	var worksheet = ctx.workbook.worksheets.add(wSheetName);
	worksheet.load('name');
	return ctx.sync().then(function() {
		console.log(worksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getActiveWorksheet()
取得活頁簿中目前作用中的工作表。

#### 語法
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### 參數
無

#### 會傳回
[Worksheet](worksheet.md)

#### 範例

```js
Excel.run(function (ctx) {  
	var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
	activeWorksheet.load('name');
	return ctx.sync().then(function() {
			console.log(activeWorksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItem(key: string)
使用其名稱或 ID 取得 worksheet 物件。

#### 語法
```js
worksheetCollectionObject.getItem(key);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|索引鍵|string|工作表的名稱或 ID。|

#### 傳回
[Worksheet](worksheet.md)
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
	var worksheets = ctx.workbook.worksheets;
	worksheets.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < worksheets.items.length; i++)
		{
			console.log(worksheets.items[i].name);
			console.log(worksheets.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

