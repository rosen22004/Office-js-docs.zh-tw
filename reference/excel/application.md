# <a name="application-object-javascript-api-for-excel"></a>Application 物件 (適用於 Excel 的 JavaScript API)

代表管理活頁簿的 Excel 應用程式。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|calculationMode|string|傳回活頁簿中使用的計算模式。唯讀。可能的值為：`Automatic` Excel 會控制重新計算；`AutomaticExceptTables` Excel 會控制重新計算，但會忽略表格中的變更；`Manual`計算會在使用者要求時完成。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|重新計算 Excel 中所有目前開啟的活頁簿。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[suspendCalculationUntilNextSync()](#suspendcalculationuntilnextsync)|void|暫止計算直至呼叫下一個 "context.sync()"。一旦設定，開發人員便有責任重新計算活頁簿，以確保能夠傳播任何相依性。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="calculatecalculationtype-string"></a>calculate(calculationType: string)
重新計算 Excel 中所有目前開啟的活頁簿。

#### <a name="syntax"></a>語法
```js
applicationObject.calculate(calculationType);
```

#### <a name="parameters"></a>參數
| 參數	       | 類型    |描述|
|:---------------|:--------|:----------|:---|
|calculationType|string|指定要使用的計算類型。可能的值為：`Recalculate` 會重新計算 Excel 標示為「已變更」的所有儲存格 (即動態或已變更資料的相依)，和以程式設計方式標示為「已變更」的儲存格。`Full`這會將所有儲存格標示為「已變更」並重新計算。`FullRebuild`這會強制重建整個計算鏈結，將所有儲存格標示為「已變更」並重新計算。|

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

### <a name="suspendcalculationuntilnextsync"></a>suspendCalculationUntilNextSync()
暫止計算直至呼叫下一個 "context.sync()"。一旦設定，開發人員便有責任重新計算活頁簿，以確保能夠傳播任何相依性。

#### <a name="syntax"></a>語法
```js
applicationObject.suspendCalculationUntilNextSync();
```

#### <a name="parameters"></a>參數
無

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

