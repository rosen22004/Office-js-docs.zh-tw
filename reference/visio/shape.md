# <a name="shape-object-javascript-api-for-visio"></a>圖形物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_
>**附註：**Visio JavaScript API 目前在預覽或生產環境中不提供使用。

代表圖形類別。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|id|int|圖形的識別碼。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-id)|
|name|string|圖形的名稱。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-name)|
|選取|bool|如圖形已選取，則傳回 True。使用者可以設定 True 以明確選取圖形。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-select)|
|text|string|圖形的文字。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-text)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|超連結|[HyperlinkCollection](hyperlinkcollection.md)|傳回 Shape 物件的超連結集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-hyperlinks)|
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|傳回圖形的資料區段。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-shapeDataItems)|
|subShapes|[ShapeCollection](shapecollection.md)|取得子圖形集合。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-subShapes)|
|檢視|[ShapeView](shapeview.md)|傳回圖形的檢視。唯讀。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-view)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-load)|

## <a name="method-details"></a>方法詳細資料

### <a name="loadparam-object"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型   |描述|
|:---------------|:--------|:----------|:---|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Sample Name";
    var shape = activePage.shapes.getItem(shapeName);
    shape.load();
    return ctx.sync().then(function () {
        console.log(shape.name );
        console.log(shape.id );
        console.log(shape.Text );
        console.log(shape.Select );
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
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    shape.view.highlight = { color: "#E7E7E7", width: 100 };
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```