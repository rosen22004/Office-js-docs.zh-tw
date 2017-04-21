# <a name="chartseriescollection-object-javascript-api-for-excel"></a>ChartSeriesCollection 物件 (適用於 Excel 的 JavaScript API)

代表圖表數列的集合。

## <a name="properties"></a>屬性

| 屬性       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|Count|int|傳回集合中的數列數目。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartSeries[]](chartseries.md)|ChartSeries 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|Int|傳回集合中的數列數目。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|根據集合中的位置，擷取數列|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getcount"></a>getCount()
傳回集合中的數列數目。

#### <a name="syntax"></a>語法
```js
chartSeriesCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitematindex-number"></a>getItemAt(index: number)
根據集合中的位置，擷取數列

#### <a name="syntax"></a>語法
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[ChartSeries](chartseries.md)

#### <a name="examples"></a>範例

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

### <a name="property-access-examples"></a>屬性存取範例
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

