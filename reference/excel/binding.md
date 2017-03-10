# <a name="binding-object-javascript-api-for-excel"></a>Binding 物件 (適用於 Excel 的 JavaScript API)

代表活頁簿中定義的 Office.js 繫結。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|id|string|代表繫結識別碼。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|傳回繫結的類型。唯讀。可能的值為：Range、Table、Text。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|刪除繫結。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|傳回繫結所代表的範圍。如果繫結不是正確的類型，則會擲回錯誤。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTable()](#gettable)|[Table](table.md)|傳回繫結所代表的表格。如果繫結不是正確的類型，則會擲回錯誤。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getText()](#gettext)|string|傳回繫結所代表的文字。如果繫結不是正確的類型，則會擲回錯誤。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="delete"></a>delete()
刪除繫結。

#### <a name="syntax"></a>語法
```js
bindingObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

### <a name="getrange"></a>getRange()
傳回繫結所代表的範圍。如果繫結不是正確的類型，則會擲回錯誤。

#### <a name="syntax"></a>語法
```js
bindingObject.getRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
下列範例會使用 binding 物件，以取得相關的範圍。

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var range = binding.getRange();
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettable"></a>getTable()
傳回繫結所代表的表格。如果繫結不是正確的類型，則會擲回錯誤。

#### <a name="syntax"></a>語法
```js
bindingObject.getTable();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Table](table.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var table = binding.getTable();
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


### <a name="gettext"></a>getText()
傳回繫結所代表的文字。如果繫結不是正確的類型，則會擲回錯誤。

#### <a name="syntax"></a>語法
```js
bindingObject.getText();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
字串

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var text = binding.getText();
    binding.load('text');
    return ctx.sync().then(function() {
        console.log(text);
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
    var binding = ctx.workbook.bindings.getItemAt(0);
    binding.load('type');
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
