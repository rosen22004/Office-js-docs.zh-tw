# <a name="bindingcollection-object-javascript-api-for-excel"></a>BindingCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於活頁簿一部份的所有 binding 物件集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|count|int|傳回集合中的繫結數目。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Binding[]](binding.md)|Binding 物件的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[add(range:Range or string, bindingType: string, id: string)](#addrange-range-or-string-bindingtype-string-id-string)|[Binding](binding.md)|將新的繫結新增至特定範圍。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromNamedItem(name: string, bindingType: string, id: string)](#addfromnameditemname-string-bindingtype-string-id-string)|[Binding](binding.md)|根據活頁簿中具名的項目，新增新的繫結。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromSelection(bindingType: string, id: string)](#addfromselectionbindingtype-string-id-string)|[Binding](binding.md)|根據目前的選取範圍，新增新的繫結。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|依 ID 取得 binding 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|根據繫結在項目陣列中的位置，取得 binding 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(id: string)](#getitemornullid-string)|[Binding](binding.md)|依 ID 取得 binding 物件。如果 binding 物件不存在，傳回物件的 isNull 屬性為 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="addrange-range-or-string-bindingtype-string-id-string"></a>add(range:Range or string, bindingType: string, id: string)
將新的繫結新增至特定範圍。

#### <a name="syntax"></a>語法
```js
bindingCollectionObject.add(range, bindingType, id);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|range|Range 或 string|繫結所繫結到的範圍。可以為 Excel Range 物件或 string。如果為 string，則必須包含完整位置，包含工作表名稱|
|bindingType|string|繫結的類型。可能的值為：Range、Table、Text|
|id|string|繫結的名稱。|

#### <a name="returns"></a>傳回
[Binding](binding.md)

### <a name="addfromnameditemname-string-bindingtype-string-id-string"></a>addFromNamedItem(name: string, bindingType: string, id: string)
根據活頁簿中具名的項目，新增新的繫結。

#### <a name="syntax"></a>語法
```js
bindingCollectionObject.addFromNamedItem(name, bindingType, id);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|Name|string|建立繫結的名稱。|
|bindingType|string|繫結的類型。可能的值為：Range、Table、Text|
|id|string|繫結的名稱。|

#### <a name="returns"></a>傳回
[Binding](binding.md)

### <a name="addfromselectionbindingtype-string-id-string"></a>addFromSelection(bindingType: string, id: string)
根據目前的選取範圍，新增新的繫結。

#### <a name="syntax"></a>語法
```js
bindingCollectionObject.addFromSelection(bindingType, id);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|bindingType|string|繫結的類型。可能的值為：Range、Table、Text|
|id|string|繫結的名稱。|

#### <a name="returns"></a>傳回
[Binding](binding.md)

### <a name="getitemid-string"></a>getItem(id: string)
依 ID 取得 binding 物件。

#### <a name="syntax"></a>語法
```js
bindingCollectionObject.getItem(id);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|id|string|要擷取之 binding 物件的 ID。|

#### <a name="returns"></a>傳回
[Binding](binding.md)

#### <a name="examples"></a>範例

建立表格繫結，以監視表格中的資料變更。當資料變更時，表格的背景色彩將會變更為橙色。

```js
function addEventHandler() {
    //Create Table1
Excel.run(function (ctx) { 
    ctx.workbook.tables.add("Sheet1!A1:C4", true);
    return ctx.sync().then(function() {
             console.log("My Diet Data Inserted!");
    })
    .catch(function (error) {
             console.log(JSON.stringify(error));
    });
});
    //Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
    else {
        // If succeeded, then add event handler to the table binding.
        Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
    }
});
}
    
// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
    // highlight the table in orange to indicate data has been changed.
    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
    return ctx.sync().then(function() {
            console.log("The value in this table got changed!");
    })
    .catch(function (error) {
            console.log(JSON.stringify(error));
    });
});
}

```



#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
    return ctx.sync().then(function() {
            console.log(binding.type); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitematindex-number"></a>getItemAt(index: number)
根據繫結在項目陣列中的位置，取得 binding 物件。

#### <a name="syntax"></a>語法
```js
bindingCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### <a name="returns"></a>傳回
[Binding](binding.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
    return ctx.sync().then(function() {
            console.log(binding.type); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemornullid-string"></a>getItemOrNull(id: string)
依 ID 取得 binding 物件。如果 binding 物件不存在，傳回物件的 isNull 屬性為 true。

#### <a name="syntax"></a>語法
```js
bindingCollectionObject.getItemOrNull(id);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|id|string|要擷取之 binding 物件的 ID。|

#### <a name="returns"></a>傳回
[Binding](binding.md)

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
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
取得繫結的數目

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('count');
    return ctx.sync().then(function() {
        console.log("Bindings: Count= " + bindings.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
