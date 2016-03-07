# RangeBorderCollection 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表構成範圍框線的 border 物件。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Count|int|集合中的 border 物件數目。唯讀。|
|items|[RangeBorder[]](rangeborder.md)|RangeBorder 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[getItem(index: string)](#getitemindex-string)|[RangeBorder](rangeborder.md)|使用名稱取得 border 物件。|
|[getItemAt(index: number)](#getitematindex-number)|[RangeBorder](rangeborder.md)|使用索引取得 border 物件。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### getItem(index: string)
使用名稱取得 border 物件。 

#### 語法
```js
rangeBorderCollectionObject.getItem(index);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|index|string|要擷取之 border 物件的索引值。可能的值為：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight、InsideVertical、InsideHorizontal、DiagonalDown、DiagonalUp。|

#### 傳回
[RangeBorder](rangeborder.md)

#### 範例
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var borderName = 'EdgeTop';
	var border = range.format.borders.getItem(borderName);
	border.load('style');
	return ctx.sync().then(function() {
			console.log(border.style);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


#### 範例
```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var border = ctx.workbook.borders.getItemAt(0);
	border.load('sideIndex');
	return ctx.sync().then(function() {
			console.log(border.sideIndex);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItemAt(index: number)
使用索引取得 border 物件。

#### 語法
```js
rangeBorderCollectionObject.getItemAt(index);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[RangeBorder](rangeborder.md)

#### 範例
```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var border = ctx.workbook.borders.getItemAt(0);
	border.load('sideIndex');
	return ctx.sync().then(function() {
			console.log(border.sideIndex);
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
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var borders = range.format.borders;
	border.load('items');
	return ctx.sync().then(function() {
		console.log(borders.count);
		for (var i = 0; i < borders.items.length; i++)
		{
			console.log(borders.items[i].sideIndex);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
以下範例會在範圍周圍增加格線框線。

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
	range.format.borders.getItem('InsideVertical').style = 'Continuous';
	range.format.borders.getItem('EdgeBottom').style = 'Continuous';
	range.format.borders.getItem('EdgeLeft').style = 'Continuous';
	range.format.borders.getItem('EdgeRight').style = 'Continuous';
	range.format.borders.getItem('EdgeTop').style = 'Continuous';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
