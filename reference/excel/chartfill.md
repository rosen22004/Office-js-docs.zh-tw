# <a name="chartfill-object-javascript-api-for-excel"></a>ChartFill 物件 (適用於 Excel 的 JavaScript API)

代表圖表項目的填滿格式。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|清除圖表項目的填滿色彩。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|將圖表項目的填滿色彩設定為統一的顏色。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="clear"></a>clear()
清除圖表項目的填滿色彩。

#### <a name="syntax"></a>語法
```js
chartFillObject.clear();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

清除名為 "Chart1" 之圖表的數值軸上主要格線的線條格式

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;    
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

### <a name="setsolidcolorcolor-string"></a>setSolidColor(color: string)
將圖表項目的填滿色彩設定為統一的顏色。

#### <a name="syntax"></a>語法
```js
chartFillObject.setSolidColor(color);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Color|string|代表框線色彩的 HTML 色彩代碼，顯示為 #RRGGBB 格式 (例如 "FFA500") 或具名 HTML 色彩 (例如 "orange")。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

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
