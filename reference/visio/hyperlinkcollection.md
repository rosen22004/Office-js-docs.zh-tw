# <a name="hyperlinkcollection-object-javascript-api-for-visio"></a>HyperlinkCollection 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_

代表超連結集合。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述|
|:---------------|:--------|:----------|
|項目|[Hyperlink[]](hyperlink.md)|超連結物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|Int|取得超連結數目。|
|[getItem(Key: number 或 string)](#getitemkey-number-or-string)|[超連結](hyperlink.md)|取得使用其機碼 (名稱或 ID) 的超連結。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


### <a name="getcount"></a>getCount()
取得超連結數目。

#### <a name="syntax"></a>語法
```js
hyperlinkCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
取得使用其機碼 (名稱或 ID) 的超連結。

#### <a name="syntax"></a>語法
```js
hyperlinkCollectionObject.getItem(Key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|機碼|number 或 string|機碼是要擷取之超連結的名稱或索引。|

#### <a name="returns"></a>傳回
[超連結](hyperlink.md)

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
    var shapeName = "Manager Belt";
    var shape = activePage.shapes.getItem(shapeName);
    var hyperlinks = shape.hyperlinks;
    shapeHyperlinks.load();
        ctx.sync().then(function () {
            for(var i=0; i<shapeHyperlinks.items.length;i++)
                {
                  var hyperlink = shapeHyperlinks.items[i];
                  console.log("Description:"+hyperlink.description +"Address:"+hyperlink.address +"SubAddress:  "+ hyperlink.subAddress);
                }

            });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
