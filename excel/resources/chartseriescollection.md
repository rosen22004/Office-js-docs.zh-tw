# ChartSeriesCollection 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表圖表數列的集合。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Count|int|傳回集合中的數列數目。唯讀。|
|items|[ChartSeries[]](chartseries.md)|ChartSeries 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|根據集合中的位置，擷取數列|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### getItemAt(index: number)
根據集合中的位置，擷取數列。

#### 語法
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[ChartSeries](chartseries.md)

#### 範例

取得數列集合中第一個數列的名稱。

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('items');
	return ctx.sync().then(function() {
		console.log(seriesCollection.items[0].name);
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
取得數列集合中的數列名稱。

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < seriesCollection.items.length; i++)
		{
			console.log(seriesCollection.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

取得集合中圖表數列的數目。

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('count');
	return ctx.sync().then(function() {
		console.log("series: Count= " + seriesCollection.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


