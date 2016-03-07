# ChartFill 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表圖表項目的填滿格式。

## 屬性

無

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除圖表項目的填滿色彩。|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|將圖表項目的填滿色彩設定為統一的顏色。|

## 方法詳細資料

### clear()
清除圖表項目的填滿色彩。

#### 語法
```js
chartFillObject.clear();
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
### setSolidColor(color: string)
將圖表項目的填滿色彩設定為統一的顏色。

#### 語法
```js
chartFillObject.setSolidColor(color);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|Color|string|代表框線色彩的 HTML 色彩代碼，顯示為 #RRGGBB 格式 (例如 "FFA500") 或具名 HTML 色彩 (例如 "orange")。|

#### 傳回
void

#### 範例

將 Chart1 的 backGround 色彩設定為紅色。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

	chart.format.fill.setSolidColor("#FF0000");

	return ctx.sync().then(function() {
			console.log("Chart1 Background Color Changed.");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

