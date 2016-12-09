# <a name="tablecolumncollection-object-javascript-api-for-excel"></a>TableColumnCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於表格一部份的所有欄集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|count|int|傳回表格中的欄數。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[TableColumn[]](tablecolumn.md)|TableColumn 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean 或 string 或 number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|將新的欄加入至表格中。|[1.1，1.1 需要索引小於總計欄數；1.4 則允許索引為選用 (null 或 -1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number 或 string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|依名稱或 ID 取得 column 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|根據資料行在集合中的位置，取得資料行。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: number or string)](#getitemornullkey-number-or-string)|[TableColumn](tablecolumn.md)|依名稱或 ID 取得 column 物件。如果資料行物件不存在，傳回物件的 isNull 屬性為 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addindex-number-values-boolean-or-string-or-number"></a>add(index: number, values: (boolean or string or number)[][])
將新的欄加入至表格中。

#### <a name="syntax"></a>語法
```js
tableColumnCollectionObject.add(index, values);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|index|number|選用。指定新欄的相對位置。如果是 null 或 -1，則會加入至結尾處。索引較高的欄將會移至側邊。以 0 開始編製索引。|
|values|(boolean or string or number)[][]|選用。表格欄中未格式化值的 2 維陣列。|

#### <a name="returns"></a>傳回
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = tables.getItem("Table1").columns.add(null, values);
    column.load('name');
    return ctx.sync().then(function() {
        console.log(column.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
依名稱或 ID 取得 column 物件。

#### <a name="syntax"></a>語法
```js
tableColumnCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|key|number 或 string| 欄名稱或 ID。|

#### <a name="returns"></a>傳回
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem('Table1').columns.getItem(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
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
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitematindex-number"></a>getItemAt(index: number)
根據欄在集合中的位置，取得欄。

#### <a name="syntax"></a>語法
```js
tableColumnCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemornullkey-number-or-string"></a>getItemOrNull(key: number or string)
依名稱或 ID 取得 column 物件。如果資料行物件不存在，傳回物件的 isNull 屬性為 true。

#### <a name="syntax"></a>語法
```js
tableColumnCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|key|number 或 string| 欄名稱或 ID。|

#### <a name="returns"></a>傳回
[TableColumn](tablecolumn.md)

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
    var tablecolumns = ctx.workbook.tables.getItem('Table1').columns;
    tablecolumns.load('items');
    return ctx.sync().then(function() {
        console.log("tablecolumns Count: " + tablecolumns.count);
        for (var i = 0; i < tablecolumns.items.length; i++)
        {
            console.log(tablecolumns.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```