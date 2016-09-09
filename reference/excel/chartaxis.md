# ChartAxis 物件 (適用於 Excel 的 JavaScript API)

代表圖表中的單個座標軸。

## 屬性

| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|majorUnit|物件|代表兩個主要刻度標記之間的間隔。可以設定為數值或空字串。傳回值一律為數字。|
|maximum|物件|代表數值軸的最大值。可以設定為數值或空字串 (針對自動數值軸)。傳回值一律為數字。|
|minimum|物件|代表數值軸的最小值。可以設定為數值或空字串 (針對自動數值軸)。傳回值一律為數字。|
|minorUnit|物件|代表兩個次要刻度標記之間的間隔。可以設定為數值或空字串 (針對自動數值軸)。傳回值一律為數字。|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
| 關聯性 | 類型	   |說明|
|:---------------|:--------|:----------|
|format|[ChartAxisFormat](chartaxisformat.md)|代表 chart 物件的格式，其中包含線條和字型格式。唯讀。|
|majorGridlines|[ChartGridlines](chartgridlines.md)|傳回 Gridlines 物件，該物件代表指定座標軸的主要格線。唯讀。|
|minorGridlines|[ChartGridlines](chartgridlines.md)|傳回 Gridlines 物件，該物件代表指定座標軸的次要格線。唯讀。|
|職稱|[ChartAxisTitle](chartaxistitle.md)|代表座標軸標題。唯讀。|

## 方法

| 方法           | 傳回類型    |說明|
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
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void
### 屬性存取範例
從 Chart1 取得圖表座標軸的 `maximum`。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var axis = chart.axes.valueaxis;
    axis.load('maximum');
    return ctx.sync().then(function() {
            console.log(axis.maximum);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

設定數值軸的 `maximum`、`minimum`、`majorunit` 或 `minorunit`。 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.axes.valueaxis.maximum = 5;
    chart.axes.valueaxis.minimum = 0;
    chart.axes.valueaxis.majorunit = 1;
    chart.axes.valueaxis.minorunit = 0.2;
    return ctx.sync().then(function() {
            console.log("Axis Settings Changed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
