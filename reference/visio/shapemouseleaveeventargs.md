# <a name="shapemouseleaveeventargs-object-javascript-api-for-visio"></a>ShapeMouseLeaveEventArgs 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_

提供引發 MouseLeave 事件之圖形的相關資訊。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述
|:---------------|:--------|:----------|
|shapeName|string|取得 MouseLeave 事件之圖形物件的名稱。|
|pageName|string|取得具有引發 MouseLeave 事件的圖形物件的頁面名稱。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無

## <a name="methods"></a>方法
無

### <a name="property-access-examples"></a>屬性存取範例
```js
Visio.run(function (ctx) { 
  var document1= ctx.document;
               var page = document1.getActivePage();
    eventResult2 = document1.onMouseLeave.add(
                function (args){            
                         console.log(Date.now()+":OnMouseLeave Event"+JSON.stringify(args));
            });
    return ctx.sync().then(function () {
           console.log("Success");
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```