# <a name="nameditemcollection-object-(javascript-api-for-excel)"></a>NamedItemCollection 物件 (適用於 Excel 的 JavaScript API)

屬於活頁簿一部份的所有 NamedItem 物件的集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|items|[NamedItem[]](nameditem.md)|NamedItem 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|使用名稱取得 nameditem 物件。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


### <a name="getitem(name:-string)"></a>getItem(name: string)
使用名稱取得 nameditem 物件。

#### <a name="syntax"></a>語法
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|Name|string|nameditem 名稱。|

#### <a name="returns"></a>傳回
[NamedItem](nameditem.md)

#### <a name="examples"></a>範例

```js
Excel.run(function (ctx) { 
    var nameditem = ctx.workbook.names.getItem(wSheetName);
    nameditem.load('type');
    return ctx.sync().then(function() {
            console.log(nameditem.type);
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
    var nameditem = ctx.workbook.names.getItemAt(0);
    nameditem.load('name');
    return ctx.sync().then(function() {
            console.log(nameditem.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="load(param:-object)"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < nameditems.items.length; i++)
        {
            console.log(nameditems.items[i].name);
            console.log(nameditems.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

取得具名項目的數目。

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('count');
    return ctx.sync().then(function() {
        console.log("nameditems: Count= " + nameditems.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
