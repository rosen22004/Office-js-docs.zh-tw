# <a name="table-object-javascript-api-for-excel"></a>Table 物件 (適用於 Excel 的 JavaScript API)

表示 Excel 表格。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|highlightFirstColumn|bool|指出第一個資料行是否包含特殊格式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|highlightLastColumn|bool|指出最後一個資料行是否包含特殊格式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|id|int|傳回可唯一識別特定活頁簿中表格的值。即使將表格重新命名，識別碼的值仍保持不變。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|表格的名稱。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedColumns|bool|表示資料行是否顯示帶狀格式，其中奇數的資料行會以不同於偶數資料行的方式反白顯示，讓閱讀資料表更方便。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedRows|bool|表示資料列是否顯示帶狀格式，其中奇數的資料列會以不同於偶數資料列的方式反白顯示，讓閱讀資料表更方便。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showFilterButton|bool|表示篩選按鈕是否在各個資料行標頭上方可見。只有在資料表包含標頭資料列時允許設定這個選項。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showHeaders|bool|指出是否顯示標題列。可以設定此值，以顯示或移除標題列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTotals|bool|指出是否顯示合計列。可以設定此值，以顯示或移除合計列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|代表表格樣式的常數值。可能的值為：TableStyleLight1 到 TableStyleLight21、TableStyleMedium1 到 TableStyleMedium28、TableStyleStyleDark1 到 TableStyleStyleDark11。也可以指定在活頁簿中呈現自訂的使用者定義樣式。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|columns|[TableColumnCollection](tablecolumncollection.md)|傳回表格中所有欄的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rows|[TableRowCollection](tablerowcollection.md)|傳回表格中所有列的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[TableSort](tablesort.md)|代表資料表的排序。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|包含目前資料表的工作表。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|清除目前在資料表上套用的所有篩選器。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToRange()](#converttorange)|[Range](range.md)|將資料表轉換成一般儲存格範圍。所有的資料會保留。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|刪除表格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|取得與表格的資料主體相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|取得與表格的標題列相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|取得與整個表格相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|取得與表格的合計列相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapplyFilters()](#reapplyfilters)|void|重新套用目前在資料表上的所有篩選器。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="clearfilters"></a>clearFilters()
清除目前在資料表上套用的所有篩選器。

#### <a name="syntax"></a>語法
```js
tableObject.clearFilters();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

### <a name="converttorange"></a>convertToRange()
將資料表轉換成一般儲存格範圍。所有的資料會保留。

#### <a name="syntax"></a>語法
```js
tableObject.convertToRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="delete"></a>delete()
刪除表格。

#### <a name="syntax"></a>語法
```js
tableObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange"></a>getDataBodyRange()
取得與表格的資料主體相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableObject.getDataBodyRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableDataRange = table.getDataBodyRange();
    tableDataRange.load('address')
    return ctx.sync().then(function() {
            console.log(tableDataRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getheaderrowrange"></a>getHeaderRowRange()
取得與表格的標題列相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableObject.getHeaderRowRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('address');
    return ctx.sync().then(function() {
        console.log(tableHeaderRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrange"></a>getRange()
取得與整個表格相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableObject.getRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load('address'); 
    return ctx.sync().then(function() {
            console.log(tableRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange"></a>getTotalRowRange()
取得與表格的合計列相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableObject.getTotalRowRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableTotalsRange = table.getTotalRowRange();
    tableTotalsRange.load('address');   
    return ctx.sync().then(function() {
            console.log(tableTotalsRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


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

### <a name="reapplyfilters"></a>reapplyFilters()
重新套用目前在資料表上的所有篩選器。

#### <a name="syntax"></a>語法
```js
tableObject.reapplyFilters();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例

依名稱取得表格。 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('index')
    return ctx.sync().then(function() {
            console.log(table.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

依索引取得表格。

```js
Excel.run(function (ctx) { 
    var index = 0;
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('id')
    return ctx.sync().then(function() {
            console.log(table.id);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

設定表格樣式。 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.name = 'Table1-Renamed';
    table.showTotals = false;
    table.style = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.style);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
