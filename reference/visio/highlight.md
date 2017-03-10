# <a name="highlight-object-javascript-api-for-visio"></a>Highlight 物件 (適用於 Visio 的 Javascript)

適用於：_Visio Online_

代表新增至圖形的螢光筆資料

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述|
|:---------------|:--------|:----------|
|Color|string|指定螢光筆顏色的字串。它必須具有表格 "#RRGGBB"，每個字母均代表介於 0 到 F 的十六進位數字，其中 RR 為介於 0 到 0xFF (255) 的紅色值，GG 為介於 0 到 0xFF (255) 的綠色值，BB 為介於 0 到 0xFF (255) 的藍色值。|
|寬度|Int|指定螢光筆筆劃寬度，以像素為單位的正整數。|

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
    shape.view.highlight = { color: "#E7E7E7", width: 100 };
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
