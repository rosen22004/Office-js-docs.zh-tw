# ChartSeries 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表圖表中的數列。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Name|string|代表圖表中的數列名稱。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
| 關聯性 | 類型	|說明|
|:---------------|:--------|:----------|
|format|[ChartSeriesFormat](chartseriesformat.md)|代表圖表數列的格式設定，其中包含填滿和線條格式。唯讀。|
|points|[ChartPointsCollection](chartpointscollection.md)|代表數列中所有點的集合。唯讀。|

## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

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

將 Chart1 的第一個數列重新命名為 "New Series Name"。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.series.getItemAt(0).name = "New Series Name";
	return ctx.sync().then(function() {
			console.log("Series1 Renamed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

