# <a name="application-object-(javascript-api-for-excel)"></a>應用程式物件 (適用於 Excel 的 JavaScript API)

代表管理活頁簿的 Excel 應用程式。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|calculationMode|string|傳回活頁簿中使用的計算模式。唯讀。可能的值為：`Automatic` Excel 控制重新計算；`AutomaticExceptTables` Excel 控制重新計算，但忽略資料表中的變更；`Manual` 當使用者要求時完成計算。|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|重新計算 Excel 中所有目前開啟的活頁簿。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## <a name="method-details"></a>方法詳細資料


### <a name="calculate(calculationtype:-string)"></a>calculate(calculationType: string)
重新計算 Excel 中所有目前開啟的活頁簿。

#### <a name="syntax"></a>語法
```js
applicationObject.calculate(calculationType);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|calculationType|string|指定要使用的計算類型。可能的值為：`Recalculate` 預設選項，藉由計算活頁簿中的所有公式來執行一般計算；`Full` 強制完整計算資料；`FullRebuild` 強制完整計算資料，並重建相依性。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    ctx.workbook.application.calculate('Full');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="load(param:-object)"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關係名稱，做為分隔字串或陣列。或者接受 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例
```js
Excel.run(function (ctx) { 
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

