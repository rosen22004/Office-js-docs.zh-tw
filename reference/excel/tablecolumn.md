# <a name="tablecolumn-object-javascript-api-for-excel"></a>TableColumn 物件 (適用於 Excel 的 JavaScript API)

代表表格中的一欄。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|id|Int|傳回可在表格內識別欄的唯一索引鍵。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|index|Int|傳回表格欄集合內欄的索引編號。以 0 開始編製索引。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|代表表格欄的名稱。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|代表所指定範圍的原始值。傳回的資料可能是 string、number 或 boolean 類型。包含錯誤的儲存格會傳回錯誤字串。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|篩選|[Filter](filter.md)|擷取套用至資料行的篩選器。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|從表格中刪除欄。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|取得與欄的資料主體相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|取得與欄的標題列相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|取得與整個欄相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|取得與欄的合計列相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="delete"></a>delete()
從表格中刪除欄。

#### <a name="syntax"></a>語法
```js
tableColumnObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
    column.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange"></a>getDataBodyRange()
取得與欄的資料主體相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableColumnObject.getDataBodyRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var dataBodyRange = column.getDataBodyRange();
    dataBodyRange.load('address');
    return ctx.sync().then(function() {
        console.log(dataBodyRange.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getheaderrowrange"></a>getHeaderRowRange()
取得與欄的標題列相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableColumnObject.getHeaderRowRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var headerRowRange = columns.getHeaderRowRange();
    headerRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(headerRowRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getrange"></a>getRange()
取得與整個欄相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableColumnObject.getRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var columnRange = columns.getRange();
    columnRange.load('address');
    return ctx.sync().then(function() {
        console.log(columnRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange"></a>getTotalRowRange()
取得與欄的合計列相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableColumnObject.getTotalRowRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).columns.getItemAt(0);
    var totalRowRange = columns.getTotalRowRange();
    totalRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(totalRowRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>屬性存取範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).columns.getItem(0);
    column.load('index');
    return ctx.sync().then(function() {
        console.log(column.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var tables = ctx.workbook.tables;
    var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = ctx.workbook.tables.getItem(tableName).columns.getItemAt(2);
    column.values = newValues;
    column.load('values');
    return ctx.sync().then(function() {
        console.log(column.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```