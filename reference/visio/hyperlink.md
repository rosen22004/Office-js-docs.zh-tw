# <a name="hyperlink-object-javascript-api-for-visio"></a>超連結物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_

代表超連結。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述|
|:---------------|:--------|:----------|
|地址|string|取得超連結物件的位址。唯讀。|
|描述|string|取得超連結的描述。唯讀。|
|subAddress|string|取得超連結物件的子位址。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


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
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    var hyperlink = shape.hyperlinks.getItem(0);
    hyperlink.load();
    return ctx.sync().then(function() {
        console.log(hyperlink.description);
        console.log(hyperlink.address);
        console.log(hyperlink.subAddress);
     });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
