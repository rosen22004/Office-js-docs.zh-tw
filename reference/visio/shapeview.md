# <a name="shapeview-object-javascript-api-for-visio"></a>ShapeView 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_
>**附註：**Visio JavaScript API 目前是預覽模式，可能有所異動。Visio JavaScript API 目前不支援在生產環境中使用。

代表 ShapeView 類別。

## <a name="properties"></a>屬性

無

## <a name="relationships"></a>關聯性
無

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 意見反應|
|:---------------|:--------|:----------|:---|
|[addOverlay(OverlayType:OverlayType、內容：字串、HorizontalAlignment︰HorizontalAlignment、VerticalAlignment：VerticalAlignment、寬度：數字、高度：數字)](#addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number)|int|在圖形頂部新增重疊。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-addOverlay)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-load)|
|[removeOverlay(OverlayId: number)](#removeoverlayoverlayid-number)|無效|移除圖形上的特定重疊或所有重疊。|[移至](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-removeOverlay)|

## <a name="method-details"></a>方法詳細資料


### <a name="addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number"></a>addOverlay(OverlayType:OverlayType、內容：字串、HorizontalAlignment︰HorizontalAlignment、VerticalAlignment：VerticalAlignment、寬度︰數字、高度︰數字)
在圖形頂部新增重疊。

#### <a name="syntax"></a>語法
```js
shapeViewObject.addOverlay(OverlayType, Content, HorizontalAlignment, VerticalAlignment, Width, Height);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|OverlayType|OverlayType|重疊類型 - 文字、影像。|
|內容|string|重疊內容。|
|HorizontalAlignment|HorizontalAlignment|重疊水平對齊 - 靠左、置中、靠右|
|VerticalAlignment|VerticalAlignment|重疊垂直對齊 - 靠上、置中、靠右|
|寬度|數字|重疊寬度。|
|高度|數字|重疊高度。|

#### <a name="returns"></a>傳回
Int

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
無效

### <a name="removeoverlayoverlayid-number"></a>removeOverlay(OverlayId: number)
移除圖形上的特定重疊或所有重疊。

#### <a name="syntax"></a>語法
```js
shapeViewObject.removeOverlay(OverlayId);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|OverlayId|number|重疊 ID。從圖形移除特定的重疊 ID。|

#### <a name="returns"></a>傳回
void

### <a name="property-access-examples"></a>屬性存取範例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    var overlayId=shape.view.addOverlay(1, "Visio Online", 2, 2, 50, 50);
    return ctx.sync();
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
    shape.view.removeOverlay(1);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
