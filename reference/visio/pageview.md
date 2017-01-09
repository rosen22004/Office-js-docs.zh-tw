# <a name="pageview-object-javascript-api-for-visio"></a>PageView 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_
>**附註：**Visio JavaScript API 目前是預覽模式，可能有所異動。Visio JavaScript API 目前不支援在生產環境中使用。

代表 PageView 類別。

## <a name="properties"></a>屬性

| 屬性	 | 類型	 |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|縮放|int|GetSet 頁面的縮放層級。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-zoom)|

## <a name="relationships"></a>關聯性

無

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|[centerViewportOnShape(ShapeId: number)](#centerviewportonshapeshapeid-number)|無效|移動瀏覽 Visio 繪圖以將指定的圖形放在檢視的中心。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-centerViewportOnShape)|
|[fitToWindow()](#fittowindow)|無效|依目前視窗調整頁面大小。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-fitToWindow)|
|[isShapeInViewport(Shape:Shape)](#isshapeinviewportshape-shape)|bool|檢查圖形是否在頁面檢視內。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-isShapeInViewport)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-load)|

## <a name="method-details"></a>方法詳細資料


### <a name="centerviewportonshapeshapeid-number"></a>centerViewportOnShape(ShapeId: number)
移動瀏覽 Visio 繪圖以將指定的圖形放在檢視的中心。

#### <a name="syntax"></a>語法
```js
pageViewObject.centerViewportOnShape(ShapeId);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|ShapeId|數字|顯示在畫面中間的 ShapeId。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    activePage.view.centerViewportOnShape(shape.Id);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="fittowindow"></a>fitToWindow()
依目前視窗調整頁面大小。

#### <a name="syntax"></a>語法
```js
pageViewObject.fitToWindow();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

### <a name="isshapeinviewportshape-shape"></a>isShapeInViewport(Shape:圖形)
檢查圖形是否在頁面檢視內。

#### <a name="syntax"></a>語法
```js
pageViewObject.isShapeInViewport(Shape);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|Shape|Shape|要檢查的圖形。|

#### <a name="returns"></a>傳回
bool

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
無效

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|位置|位置|位置物件，在檢視內指定頁面的新位置。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    activePage.view.zoom = 300;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

