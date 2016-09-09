# BindingCollection 物件 (適用於 Excel 的 JavaScript API)

代表屬於活頁簿一部份的所有 binding 物件集合。

## 屬性

| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|Count|int|傳回集合中的繫結數目。唯讀。|
|項目|[Binding[]](binding.md)|Binding 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|依 ID 取得 binding 物件。|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|根據繫結在項目陣列中的位置，取得 binding 物件。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料


### getItem(id: string)
依 ID 取得 binding 物件。

#### 語法
```js
bindingCollectionObject.getItem(id);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|id|string|要擷取之 binding 物件的 ID。|

#### 傳回
[Binding](binding.md)

#### 範例

建立表格繫結，以監視表格中的資料變更。當資料變更時，表格的背景色彩會變更為橙色。

```js
(function () {
    // Create myTable
    Excel.run(function (ctx) {
        var table = ctx.workbook.tables.add("Sheet1!A1:C4", true);
        table.name = "myTable";
        return ctx.sync().then(function () {
            console.log("MyTable is Created!");

            //Create a new table binding for myTable
            Office.context.document.bindings.addFromNamedItemAsync("myTable", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
                if (asyncResult.status == "failed") {
                    console.log("Action failed with error: " + asyncResult.error.message);
                }
                else {
                    // If successful, add the event handler to the table binding.
                    Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
            });
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
    });
    
    // When data in the table is changed, this event is triggered.
    function onBindingDataChanged(eventArgs) {
        Excel.run(function (ctx) {
            // Highlight the table in orange to indicate data changed.
            var fill = ctx.workbook.tables.getItem("myTable").getDataBodyRange().format.fill;
            fill.load("color");
            return ctx.sync().then(function () {
                if (fill.color != "Orange") {
                    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
 
                    console.log("The value in this table got changed!");
                }
                else
                    
            })
                .then(ctx.sync)
            .catch(function (error) {
                console.log(JSON.stringify(error));
            });
        });
    } 
})();
 


```



#### 範例
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


### getItemAt(index: number)
根據繫結在項目陣列中的位置，取得 binding 物件。

#### 語法
```js
bindingCollectionObject.getItemAt(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[Binding](binding.md)

#### 範例
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


### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關係名稱，做為分隔字串或陣列。或者接受 [loadOption](loadoption.md) 物件。|

#### 傳回
void
### 屬性存取範例

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
            console.log(bindings.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
取得繫結的數目。

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
