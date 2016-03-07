# ChartLineFormat 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

封裝線條元素的格式設定選項。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Color|string|代表圖表中線條色彩的 HTML 色彩代碼。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除圖表項目的線條格式。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### clear()
清除圖表項目的線條格式。

#### 語法
```js
chartLineFormatObject.clear();
```

#### 參數
無

#### 傳回
void

#### 範例

清除名為 "Chart1" 之圖表的數值軸上主要格線的線條格式。

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	
	gridlines.format.line.clear();
	return ctx.sync().then(function() {
			console.log("Chart Major Gridlines Format Cleared");
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

將圖表數值軸的主要格線設定為紅色。

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.axes.valueaxis.majorGridlines;
	gridlines.format.line.color = "#FF0000";
	return ctx.sync().then(function() {
			console.log("Chart Gridlines Color Updated");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

