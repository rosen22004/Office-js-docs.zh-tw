# <a name="chartpointscollection-object-javascript-api-for-excel"></a>ChartPointsCollection 物件 (適用於 Excel 的 JavaScript API)

圖表內數列中所有圖表點的集合。

## <a name="properties"></a>屬性

| 屬性       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|Count|Int|傳回系列中的圖表點數目。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartPoint[]](chartpoint.md)|ChartPoints 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|Int|傳回系列中的圖表點數目。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|根據數列中的位置來擷取點。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getcount"></a>getCount()
傳回系列中的圖表點數目。

#### <a name="syntax"></a>語法
```js
chartPointsCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitematindex-number"></a>getItemAt(index: number)
根據數列中的位置來擷取點。

#### <a name="syntax"></a>語法
```js
chartPointsCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[ChartPoint](chartpoint.md)

#### <a name="examples"></a>範例
設定點集合中第一個點的框線色彩

```js
Excel.run(function (ctx) { 
    var points = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
    return ctx.sync().then(function() {
        console.log("Point Border Color Changed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```### Property access examples

Get the names of points in the points collection

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
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

取得點的數目

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
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
