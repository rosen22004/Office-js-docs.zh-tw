# <a name="worksheetcollection-object-javascript-api-for-excel"></a>WorksheetCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於活頁簿一部份的 worksheet 物件集合。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|項目|[Worksheet[]](worksheet.md)|Worksheet 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|將新的工作表加入活頁簿中。工作表會加入至現有工作表的結尾處。如果您想要啟動新加入的工作表，請對其呼叫 ".activate()。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|取得活頁簿中目前作用中的工作表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|使用其名稱或 ID 取得 worksheet 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullkey-string)|[Worksheet](worksheet.md)|使用其名稱或 ID 取得 worksheet 物件。如果工作表不存在，傳回物件的 isNull 屬性為 true。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addname-string"></a>add(name: string)
將新的工作表加入活頁簿中。工作表會加入至現有工作表的結尾處。如果您想要啟動新加入的工作表，請對其呼叫 ".activate()。

#### <a name="syntax"></a>語法
```js
worksheetCollectionObject.add(name);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|Name|string|選用。要加入的工作表的名稱。若指定，名稱應為唯一的。若不指定，Excel 會自行決定新工作表的名稱。|

#### <a name="returns"></a>傳回
[Worksheet](worksheet.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sample Name';
    var worksheet = ctx.workbook.worksheets.add(wSheetName);
    worksheet.load('name');
    return ctx.sync().then(function() {
        console.log(worksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getactiveworksheet"></a>getActiveWorksheet()
取得活頁簿中目前作用中的工作表。

#### <a name="syntax"></a>語法
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Worksheet](worksheet.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) {  
    var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(activeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemkey-string"></a>getItem(key: string)
使用其名稱或 ID 取得 worksheet 物件。

#### <a name="syntax"></a>語法
```js
worksheetCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|索引鍵|string|工作表的名稱或 ID。|

#### <a name="returns"></a>傳回
[Worksheet](worksheet.md)

### <a name="getitemornullkey-string"></a>getItemOrNull(key: string)
使用其名稱或 ID 取得 worksheet 物件。如果工作表不存在，傳回物件的 isNull 屬性為 true。

#### <a name="syntax"></a>語法
```js
worksheetCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|索引鍵|string|工作表的名稱或 ID。|

#### <a name="returns"></a>傳回
[Worksheet](worksheet.md)

### <a name="loadparam-object"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例
```js
Excel.run(function (ctx) { 
    var worksheets = ctx.workbook.worksheets;
    worksheets.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < worksheets.items.length; i++)
        {
            console.log(worksheets.items[i].name);
            console.log(worksheets.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
