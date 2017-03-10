# <a name="datarefreshcompleteeventargs-object-javascript-api-for-visio"></a>DataRefreshCompleteEventArgs 物件 (適用於 Visio 的 JavaScript API)

適用於：_Visio Online_

提供引發 DataRefreshComplete 事件的文件資訊。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述
|:---------------|:--------|:----------|
|成功|bool|取得 DataRefreshComplete 事件的 successfailure。|
|文件|[文件](document.md)|取得引發 DataRefreshComplete 事件的文件物件。|

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
         eventResult1 = document1.onDataRefreshComplete.add(
    function (args){
           console.log("Data Refresh Result: "+args.success);
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
