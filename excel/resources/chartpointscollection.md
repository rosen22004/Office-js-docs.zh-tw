# ChartPointsCollection 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

圖表內數列中所有圖表點的集合。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Count|int|傳回集合中的圖表點數目。唯讀。|
|items|[ChartPoint[]](chartpoint.md)|ChartPoints 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|根據數列中的位置來擷取點。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### getItemAt(index: number)
根據數列中的位置來擷取點。

#### 語法
```js
chartPointsCollectionObject.getItemAt(index);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[ChartPoint](chartpoint.md)

#### 範例
設定點集合中第一個點的框線色彩。

```js
Excel.run(function (ctx) { 
	var point = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
	points.getItemAt(0).format.fill.setSolidColor("#8FBC8F");
	return ctx.sync().then(function() {
		console.log("Point Border Color Changed");
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

設定點集合中的點的名稱。

```js
Excel.run(function (ctx) { 
	var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
	pointsCollection.load('items');
	return ctx.sync().then(function() {
		console.log("Points Collection loaded");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

取得點的數目。

```js
Excel.run(function (ctx) { 
	var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
	pointsCollection.load('count');
	return ctx.sync().then(function() {
		console.log("points: Count= " + pointsCollection.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

