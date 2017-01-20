# <a name="chartcollection-object-javascript-api-for-excel"></a>ChartCollection 物件 (適用於 Excel 的 JavaScript API)

工作表上所有 chart 物件的集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|Count|int|傳回工作表中的圖表數目。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Chart[]](chart.md)|Chart 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(type: string, sourceData:Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|建立新圖表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|使用其名稱取得圖表。如果有多個圖表具有相同的名稱，則會傳回第一個圖表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|根據圖表在集合中的位置，取得圖表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: string)](#getitemornullname-string)|[Chart](chart.md)|使用其名稱取得圖表。如果有多個圖表具有相同的名稱，則會傳回第一個圖表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addtype-string-sourcedata-range-seriesby-string"></a>add(type: string, sourceData:Range, seriesBy: string)
建立新圖表。

#### <a name="syntax"></a>語法
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|類型|string|代表圖表的類型。請參閱下方提供之可能適用的表格類型。|
|sourceData|Range|對應到來源資料的 Range 物件。|
|seriesBy|string|選用。指定在圖表中使用欄或列作為資料數列的方法。可能的值為：Auto、Columns、Rows|

**以下是有效的表格類型：**

`ColumnClustered`, `ColumnStacked`, `ColumnStacked100`, `_3DColumnClustered`, `_3DColumnStacked`, `_3DColumnStacked100`, `BarClustered`, `BarStacked`, `BarStacked100`, `_3DBarClustered`, `_3DBarStacked`, `_3DBarStacked100`, `LineStacked`, `LineStacked100`, `LineMarkers`, `LineMarkersStacked`, `LineMarkersStacked100`, `PieOfPie`, `PieExploded`, `_3DPieExploded`, `BarOfPie`, `XYScatterSmooth`, `XYScatterSmoothNoMarkers`, `XYScatterLines`, `XYScatterLinesNoMarkers`, `AreaStacked`, `AreaStacked100`, `_3DAreaStacked`, `_3DAreaStacked100`, `DoughnutExploded`, `RadarMarkers`, `RadarFilled`, `Surface`, `SurfaceWireframe`, `SurfaceTopView`, `SurfaceTopViewWireframe`, `Bubble`, `Bubble3DEffect`, `StockHLC`, `StockOHLC`, `StockVHLC`, `StockVOHLC`, `CylinderColClustered`, `CylinderColStacked`, `CylinderColStacked100`, `CylinderBarClustered`, `CylinderBarStacked`, `CylinderBarStacked100`, `CylinderCol`, `ConeColClustered`, `ConeColStacked`, `ConeColStacked100`, `ConeBarClustered`, `ConeBarStacked`, `ConeBarStacked100`, `ConeCol`, `PyramidColClustered`, `PyramidColStacked`, `PyramidColStacked100`, `PyramidBarClustered`, `PyramidBarStacked`, `PyramidBarStacked100`, `PyramidCol`, `_3DColumn`, `Line`, `_3DLine`, `_3DPie`, `Pie`, `XYScatter`, `_3DArea`, `Area`, `Doughnut`, `Radar`


#### <a name="returns"></a>傳回
[Chart](chart.md)

#### <a name="examples"></a>範例

使用來自範圍 "A1:B4" 的 `chartType`，在工作表 "Charts" 上新增圖表 `sourceData` "ColumnClustered"，並且 `seriresBy` 設定為 "auto"。

```js
Excel.run(function (ctx) { 
    var rangeSelection = "A1:B4";
    var range = ctx.workbook.worksheets.getItem(sheetName)
        .getRange(rangeSelection);
    var chart = ctx.workbook.worksheets.getItem(sheetName)
        .charts.add("ColumnClustered", range, "auto");  return ctx.sync().then(function() {
            console.log("New Chart Added");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemname-string"></a>getItem(name: string)
使用其名稱取得圖表。如果有多個圖表具有相同的名稱，則會傳回第一個圖表。

#### <a name="syntax"></a>語法
```js
chartCollectionObject.getItem(name);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|Name|string|要擷取之圖表的名稱。|

#### <a name="returns"></a>傳回
[Chart](chart.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var chartname = 'Chart1';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var chartId = 'SamplChartId';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
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


### <a name="getitematindex-number"></a>getItemAt(index: number)
根據圖表在集合中的位置，取得圖表。

#### <a name="syntax"></a>語法
```js
chartCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[Chart](chart.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
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


### <a name="getitemornullname-string"></a>getItemOrNull(name: string)
使用其名稱取得圖表。如果有多個圖表具有相同的名稱，則會傳回第一個圖表。

#### <a name="syntax"></a>語法
```js
chartCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|Name|string|要擷取之圖表的名稱。|

#### <a name="returns"></a>傳回
[Chart](chart.md)

### <a name="loadparam-object"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < charts.items.length; i++)
        {
            console.log(charts.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

取得圖表的數目

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('count');
    return ctx.sync().then(function() {
        console.log("charts: Count= " + charts.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

