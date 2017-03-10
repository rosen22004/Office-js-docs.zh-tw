# <a name="tablerowcollection-object-javascript-api-for-excel"></a>TableRowCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於表格一部份的所有列集合。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|Count|int|傳回表格中的列數。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[TableRow[]](tablerow.md)|TableRow 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean 或 string 或 number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|新增一或多個資料列至表格中。傳回的物件將會在新增列的頂端。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|Int|取得表格中的列數。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|根據列在集合中的位置，取得列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addindex-number-values-boolean-or-string-or-number"></a>add(index: number, values: (boolean or string or number)[][])
新增一或多個資料列至表格中。傳回的物件將會在新增資料列的頂端。

#### <a name="syntax"></a>語法
```js
tableRowCollectionObject.add(index, values);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|index|數字|選用。指定新列的相對位置。如果是 null 或 -1，則會加入至結尾處。插入列下方的任何列都會向下移。以 0 開始編製索引。|
|values|(boolean or string or number)[][]|選用。表格列中未格式化值的 2 維陣列。|

#### <a name="returns"></a>傳回
[TableRow](tablerow.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample", "Values", "For", "New", "Row"]];
    var row = tables.getItem("Table1").rows.add(null, values);
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

### <a name="getcount"></a>getCount()
取得表格中的列數。

#### <a name="syntax"></a>語法
```js
tableRowCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitematindex-number"></a>getItemAt(index: number)
根據列在集合中的位置，取得列。

#### <a name="syntax"></a>語法
```js
tableRowCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[TableRow](tablerow.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
    tablerow.load('name');
    return ctx.sync().then(function() {
            console.log(tablerow.name);
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
    var tablerows = ctx.workbook.tables.getItem('Table1').rows;
    tablerows.load('items');
    return ctx.sync().then(function() {
        console.log("tablerows Count: " + tablerows.count);
        for (var i = 0; i < tablerows.items.length; i++)
        {
            console.log(tablerows.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```