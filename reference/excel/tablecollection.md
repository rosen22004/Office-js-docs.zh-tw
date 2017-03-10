# <a name="tablecollection-object-javascript-api-for-excel"></a>TableCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於活頁簿或工作表一部份的所有表格集合，視到達方式而定。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|Count|int|傳回工作表中的表格數目。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Table[]](table.md)|Table 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(address: [object, hasHeaders: bool)](#addaddress-object-hasheaders-bool)|[表格](table.md)|建立新表格。Range 物件或來源位址將決定表格加入至哪一個工作表。如果無法加入表格 (例如因為位址無效，或是此表格會與其他表格重疊)，則將會擲回錯誤。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|Int|取得集合中的表格數目。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number 或 string)](#getitemkey-number-or-string)|[Table](table.md)|依名稱或識別碼取得資料表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|根據表格在集合中的位置，取得表格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number 或 string)](#getitemornullobjectkey-number-or-string)|[表格](table.md)|依名稱或 ID 取得表格。如果表格不存在，會傳回 null 物件。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addaddress-object-hasheaders-bool"></a>add(address: string, hasHeaders: bool)
建立新表格。Range 物件或來源位址將決定表格加入至哪一個工作表。如果無法加入表格 (例如因為位址無效，或是此表格會與其他表格重疊)，則將會擲回錯誤。

#### <a name="syntax"></a>語法
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|位址|[object|代表資料來源的 Range 物件、字串位址或範圍名稱。如果位址不含工作表名稱，則會使用目前作用中的工作表。如為 1.1，使用字串參數；如為 1.3 ，則亦可接受 Range 物件。|
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

### <a name="getcount"></a>getCount()
取得集合中的表格數目。

#### <a name="syntax"></a>語法
```js
tableCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitemkey-number-or-string"></a>getItem(key: number 或 string)
依名稱或識別碼取得資料表。

#### <a name="syntax"></a>語法
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
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
| 參數	       | 類型    |描述|
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


### <a name="getitemornullobjectkey-number-or-string"></a>getItemOrNullObject(key: number 或 string)
依名稱或 ID 取得表格。如果表格不存在，會傳回 null 物件。

#### <a name="syntax"></a>語法
```js
tableCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|key|number 或 string|要擷取之表格的名稱或 ID。|

#### <a name="returns"></a>傳回
[表格](table.md)
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