# <a name="chartaxis-object-javascript-api-for-excel"></a>ChartAxis 物件 (適用於 Excel 的 JavaScript API)

代表圖表中的單個座標軸。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|majorUnit|物件|代表兩個主要刻度標記之間的間隔。可以設定為數值或空字串。傳回值一律為數字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|maximum|物件|代表數值軸上的最大值。可以設定為數值或空字串 (針對自動數值軸)。傳回值一律為數字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minimum|物件|代表數值軸上的最小值。可以設定為數值或空字串 (針對自動數值軸)。傳回值一律為數字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minorUnit|物件|代表兩個次要刻度標記之間的間隔。可以設定為數值或空字串 (針對自動數值軸)。傳回值一律為數字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|format|[ChartAxisFormat](chartaxisformat.md)|代表 chart 物件的格式，其中包含線條和字型格式。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|majorGridlines|[ChartGridlines](chartgridlines.md)|傳回 gridlines 物件，該物件代表指定座標軸的主要格線。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minorGridlines|[ChartGridlines](chartgridlines.md)|傳回 Gridlines 物件，該物件代表指定座標軸的次要格線。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|title|[ChartAxisTitle](chartaxistitle.md)|代表座標軸標題。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
無


## <a name="method-details"></a>方法詳細資料

### <a name="property-access-examples"></a>屬性存取範例
從 Chart1 取得圖表座標軸的 `maximum`。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var axis = chart.axes.valueAxis;
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

設定數值軸的 `maximum`、`minimum`、`majorunit`、`minorunit`。 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.axes.valueAxis.maximum = 5;
    chart.axes.valueAxis.minimum = 0;
    chart.axes.valueAxis.majorUnit = 1;
    chart.axes.valueAxis.minorUnit = 0.2;
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
