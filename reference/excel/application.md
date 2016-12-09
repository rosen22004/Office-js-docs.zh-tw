# <a name="application-object-javascript-api-for-excel"></a>應用程式物件 (適用於 Excel 的 JavaScript API)

代表管理活頁簿的 Excel 應用程式。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述|需求集合|
|:---------------|:--------|:----------|:----------|
|calculationMode|string|傳回活頁簿中使用的計算模式。唯讀。可能的值為：`Automatic` Excel 控制重新計算；`AutomaticExceptTables` Excel 控制重新計算，但忽略資料表中的變更；`Manual` 當使用者要求時完成計算。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|需求集合|
|:---------------|:--------|:----------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|重新計算 Excel 中所有目前開啟的活頁簿。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="calculatecalculationtype-string"></a>calculate(calculationType: string)
重新計算 Excel 中所有目前開啟的活頁簿。

#### <a name="syntax"></a>語法
```js
applicationObject.calculate(calculationType);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|calculationType|string|指定要使用的計算類型。可能的值為：`Recalculate` 此為模糊計算且主要用於向後相容。`Full` 重新計算所有 Excel 標示為「已變更」的儲存格，也就是根據動態或已變更的資料，以及程式設計方式所標示的「已變更」。`FullRebuild` 重新計算所有開啟活頁簿的全部儲存格。|

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


### <a name="loadparam-object"></a>load(param: object)
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
