# <a name="shapedataitem-object-javascript-api-for-visio"></a>ShapeDataItem 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_
>**附註：**Visio JavaScript API 目前是預覽模式，可能有所異動。Visio JavaScript API 目前不支援在生產環境中使用。

代表 ShapeDataItem。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|標籤|string|指定圖形資料項目標籤的字串。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItem-label)|
|值|string|指定圖形資料項目值的字串。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItem-value)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeDataItem-load)|

## <a name="method-details"></a>方法詳細資料


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
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
        var shapeDataItem = shape.shapeDataItems.getItem(0);
    shapeDataItem.load();
        return ctx.sync().then(function() {
                console.log(shapeDataItem.label);
                console.log(shapeDataItem.value);
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
