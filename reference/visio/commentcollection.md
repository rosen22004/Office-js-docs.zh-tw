# <a name="commentcollection-object-javascript-api-for-visio"></a>CommentCollection 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_

代表指定圖形的 ShapeDataItemCollection。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述
|:---------------|:--------|:----------|
|項目|[Comment[]](comment.md)|註解物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|Int|取得註解數目。|
|[getItem(key: string)](#getitemkey-string)|[註解](comment.md)|使用其名稱取得註解。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


### <a name="getcount"></a>getCount()
取得註解數目。

#### <a name="syntax"></a>語法
```js
CommentCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

### <a name="getitemkey-string"></a>getItem(key: string)
使用其名稱取得註解。

#### <a name="syntax"></a>語法
```js
CommentCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|
|索引鍵|string|機碼是要擷取之註解名稱。|

#### <a name="returns"></a>傳回
[註解](comment.md)

### <a name="loadparam-object"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例
```js
 Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Position Belt.41";
    var shape = activePage.shapes.getItem(shapeName);
    var shapecomments= shape.comments;
        shapecomments.load();
        return ctx.sync().then(function () {
             for(var i=0; i<shapecomments.items.length;i++)
        {
                    var comment= shapecomments.items[i];
            console.log("comment Author: " + comment.author);
            console.log("Comment Text: " + comment.text);
            console.log("Date " + comment.date);
        }
     });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
