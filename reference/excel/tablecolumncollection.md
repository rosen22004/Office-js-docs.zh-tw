# TableColumnCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於表格一部份的所有欄集合。

## 屬性

| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|Count|int|傳回表格中的欄數。唯讀。|
|項目|[TableColumn[]](tablecolumn.md)|TableColumn 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|將新的欄加入至表格中。|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|依名稱或 ID 取得 column 物件。|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|根據欄在集合中的位置，取得欄。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料


### add(index: number, values: (boolean or string or number)[][])
將新的欄加入至表格中。

#### 語法
```js
tableColumnCollectionObject.add(index, values);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|index|number|指定新欄的相對位置。該位置的前一欄會向右移。索引值應該等於或小於最後一欄的索引值，所以不能用來將欄附加至表格結尾處。以 0 開始編製索引。|
|values|(boolean or string or number)[][]|選用。表格欄中未格式化值的 2 維陣列。|

#### 傳回
[TableColumn](tablecolumn.md)

#### 範例

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


### getItem(key: number or string)
依名稱或 ID 取得 column 物件。

#### 語法
```js
tableColumnCollectionObject.getItem(key);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|key|number 或 string| 欄名稱或 ID。|

#### 傳回
[TableColumn](tablecolumn.md)

#### 範例

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
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


#### 範例
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

### getItemAt(index: number)
根據欄在集合中的位置，取得欄。

#### 語法
```js
tableColumnCollectionObject.getItemAt(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[TableColumn](tablecolumn.md)

#### 範例
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

### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void
### 屬性存取範例

```js
Excel.run(function (ctx) { 
    var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
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
