# <a name="tablerow-object-javascript-api-for-excel"></a>TableRow 物件 (適用於 Excel 的 JavaScript API)

代表表格中的一列。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|索引|int|傳回表格列集合中列的索引編號。以 0 開始編製索引。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|代表所指定範圍的原始值。傳回的資料可能是 string、number 或 boolean 類型。包含錯誤的儲存格會傳回錯誤字串。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|從表格中刪除列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|傳回與整個列相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="delete"></a>delete()
從表格中刪除列。

#### <a name="syntax"></a>語法
```js
tableRowObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).rows.getItemAt(2);
    row.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrange"></a>getRange()
傳回與整個列相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableRowObject.getRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).rows.getItemAt(0);
    var rowRange = row.getRange();
    rowRange.load('address');
    return ctx.sync().then(function() {
        console.log(rowRange.address);
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
    var row = ctx.workbook.tables.getItem(tableName).rows.getItem(0);
    row.load('index');
    return ctx.sync().then(function() {
        console.log(row.index);
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
    var tables = ctx.workbook.tables;
    var newValues = [["New", "Values", "For", "New", "Row"]];
    var tableName = 'Table1';
    var row = ctx.workbook.tables.getItem(tableName).rows.getItemAt(2);
    row.values = newValues;
    row.load('values');
    return ctx.sync().then(function() {
        console.log(row.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```