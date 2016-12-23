# <a name="shapecollection-object-javascript-api-for-visio"></a>ShapeCollection 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_
>**附註：**Visio JavaScript API 目前是預覽模式，可能有所異動。Visio JavaScript API 目前不支援在生產環境中使用。

代表圖形集合。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|項目|[Shape[]](shape.md)|圖形物件的集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-items)|

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|Int|取得集合中的圖形數目。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getCount)|
|[getItem(key: number 或 string)](#getitemkey-number-or-string)|[圖形](shape.md)|取得使用其機碼 (名稱或索引) 的圖形。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="getcount"></a>getCount()
取得集合中的圖形數目。

#### <a name="syntax"></a>語法
```js
shapeCollectionObject.getCount();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
Int

#### <a name="examples"></a>範例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var numShapesActivePage = activePage.shapes.getCount();
    return ctx.sync().then(function () {
        console.log("Shapes Count: " + numShapesActivePage.value);
    });

}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemkey-number-or-string"></a>getItem(key: number 或 string)
取得使用其機碼 (名稱或索引) 的圖形。

#### <a name="syntax"></a>語法
```js
shapeCollectionObject.getItem(key);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|key|number 或 string|機碼是要擷取之圖形的名稱或索引。|

#### <a name="returns"></a>傳回
[圖形](shape.md)

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
