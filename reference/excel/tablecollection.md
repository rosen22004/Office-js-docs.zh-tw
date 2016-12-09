# <a name="tablecollection-object-javascript-api-for-excel"></a>TableCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於活頁簿一部份的所有表格集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|count|int|傳回工作表中的表格數目。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Table[]](table.md)|Table 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(address:Range or string, hasHeaders: bool)](#addaddress-range-or-string-hasheaders-bool)|[Table](table.md)|建立新表格。Range 物件或來源位址將決定表格加入至哪一個工作表。如果無法加入表格 (例如因為位址無效，或是此表格會與其他表格重疊)，則將會擲回錯誤。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number 或 string)](#getitemkey-number-or-string)|[Table](table.md)|依名稱或識別碼取得資料表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|根據表格在集合中的位置，取得表格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: number or string)](#getitemornullkey-number-or-string)|[Table](table.md)|依名稱或識別碼取得資料表。如果資料表不存在，傳回物件的 isNull 屬性為 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addaddress-range-or-string-hasheaders-bool"></a>add(address:Range or string, hasHeaders: bool)
建立新表格。Range 物件或來源位址將決定表格加入至哪一個工作表。如果無法加入表格 (例如因為位址無效，或是此表格會與其他表格重疊)，則將會擲回錯誤。

#### <a name="syntax"></a>語法
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|address|Range 或 string|表示資料來源的範圍的 Range 物件、字串位址或範圍名稱。如果位址不含工作表名稱，則會使用目前作用中的工作表。字串參數需要需求集合 1.1 版；接受 Range 物件則需要 1.3 版。|
|hasHeaders|bool|布林值，指出是要匯入的資料是否具有欄標籤。如果來源不含標頭 (亦即此屬性設為 false)，Excel 會自動產生標頭，並將資料向下移一列。|

#### <a name="returns"></a>傳回
[Table](table.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
    table.load('name');
    return ctx.sync().then(function() {
        console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
依名稱或識別碼取得資料表。

#### <a name="syntax"></a>語法
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|key|number 或 string|要擷取之表格的名稱或 ID。|

#### <a name="returns"></a>傳回
[Table](table.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
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
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitematindex-number"></a>getItemAt(index: number)
根據表格在集合中的位置，取得表格。

#### <a name="syntax"></a>語法
```js
tableCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[Table](table.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemornullkey-number-or-string"></a>getItemOrNull(key: number or string)
依名稱或識別碼取得資料表。如果資料表不存在，傳回物件的 isNull 屬性為 true。

#### <a name="syntax"></a>語法
```js
tableCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|key|number 或 string|要擷取之表格的名稱或 ID。|

#### <a name="returns"></a>傳回
[Table](table.md)

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
    var tables = ctx.workbook.tables;
    tables.load();
    return ctx.sync().then(function() {
        console.log("tables Count: " + tables.count);
        for (var i = 0; i < tables.items.length; i++)
        {
            console.log(tables.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

取得表格的數目

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load('count');
    return ctx.sync().then(function() {
        console.log(tables.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```