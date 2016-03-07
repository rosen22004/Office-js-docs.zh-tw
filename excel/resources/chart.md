# Chart 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Office 2016_

代表活頁簿中的 chart 物件。

## 屬性

| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|Height|double|代表 chart 物件的高度，以點為單位。|
|left|double|從圖表左側到工作表原點的距離，以點為單位。|
|name|string|代表 chart 物件的名稱。|
|top|double|代表從物件上邊緣到第 1 列頂端 (在工作表上) 或圖表區域頂端 (在圖表上) 的距離，以點為單位。|
|width|double|代表 chart 物件的寬度，以點為單位。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
| 關聯性 | 類型	|說明|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|代表圖表座標軸。唯讀。|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|代表圖表上的資料標籤。唯讀。|
|format|[ChartAreaFormat](chartareaformat.md)|封裝圖表區域的格式屬性。唯讀。|
|legend|[ChartLegend](chartlegend.md)|代表圖表的圖例。唯讀。|
|series|[ChartSeriesCollection](chartseriescollection.md)|代表圖表中的單一數列或數列集合。唯讀。|
|title|[ChartTitle](charttitle.md)|代表所指定圖表的標題，包括標題的文字、可見度、位置和格式設定。唯讀。|

## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|刪除 chart 物件。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|
|[setData(sourceData:Range or string, seriesBy: string)](#setdatasourcedata-range-or-string-seriesby-string)|void|重設圖表的來源資料。|
|[setPosition(startCell:Range or string, endCell:Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|將圖表定位至工作表上儲存格的相對位置。|

## 方法詳細資料

### delete()
刪除 chart 物件。

#### 語法
```js
chartObject.delete();
```

#### 參數
無

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.delete();
	return ctx.sync(); 
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


### setData(sourceData:Range or string, seriesBy: string)
重設圖表的來源資料。

#### 語法
```js
chartObject.setData(sourceData, seriesBy);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|sourceData|Range 或 string|包含來源資料的範圍位址或名稱。如果使用位址或工作表範圍名稱，則必須包含工作表名稱 (例如 "Sheet1!A5:B9")。 |
|seriesBy|string|選用。指定在圖表中使用欄或列作為資料數列的方法。可以是下列其中一項：Auto (預設)、Rows、Columns。可能的值為：Auto、Columns、Rows。|

#### 傳回
void

#### 範例

將 `sourceData` 設定為 "A1:B4" 並將 `seriesBy` 設定為 "Columns"。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var sourceData = "A1:B4";
	chart.setData(sourceData, "Columns");
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### setPosition(startCell:Range or string, endCell:Range or string)
將圖表定位至工作表上儲存格的相對位置。

#### 語法
```js
chartObject.setPosition(startCell, endCell);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|startCell|Range 或 string|起始儲存格。這是圖表的移動目標位置。取決於使用者的從左至右顯示設定，開始儲存格將是左上角或右上角儲存格。|
|endCell|Range 或 string|選用。結束儲存格。如果指定，則圖表的寬度和高度會設定為完全覆蓋這個儲存格/範圍。|

#### 傳回
void

#### 範例


```js
Excel.run(function (ctx) { 
	var sheetName = "Charts";
	var sourceData = sheetName + "!" + "A1:B4";
	var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
	chart.width = 500;
	chart.height = 300;
	chart.setPosition("C2", null);
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### 屬性存取範例

取得名為 "Chart1" 的圖表。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.load('name');
	return ctx.sync().then(function() {
			console.log(chart.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

更新圖表，包括重新命名、定位和調整大小。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.name="New Name";
	chart.top = 100;
	chart.left = 100;
	chart.height = 200;
	chart.weight = 200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

提供新名稱給圖表；調整圖表的大小為高度和寬度都是 200 點。將 Chart1 向左上方移動 100 點。 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
	chart.name="New Name";	
	chart.top = 100;
	chart.left = 100;
	chart.height =200;
	chart.width =200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

