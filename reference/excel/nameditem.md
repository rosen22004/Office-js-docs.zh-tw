# <a name="nameditem-object-javascript-api-for-excel"></a>NamedItem 物件 (適用於 Excel 的 JavaScript API)

代表某個儲存格範圍或值的已定義名稱。名稱可以是原始命名物件 (如下列類型所示)、範圍物件、範圍的參照。此物件可用來取得與名稱相關聯的範圍物件。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|name|string|物件的名稱。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|指出與名稱相關聯的參考類型。唯讀。可能的值為：String、Integer、Double、Boolean、Range。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|代表名稱的定義參考公式。例如 =Sheet14!$B$2:$H$12、=4.75 等。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|指定物件是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|傳回與名稱相關的範圍物件。如果具名項目的類型不是範圍，則擲回例外狀況。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="getrange"></a>getRange()
傳回與名稱相關的範圍物件。如果具名項目的類型不是範圍，則擲回例外狀況。

#### <a name="syntax"></a>語法
```js
namedItemObject.getRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例

傳回與名稱相關的範圍物件。如果名稱不是 `Range` 類型則傳回 `null`。附註：此 API 目前僅支援 Workbook 範圍項目。**

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var range = names.getItem('MyRange').getRange();
    range.load('address');
    return ctx.sync().then(function() {
            console.log(range.address);
    });
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
|:---------------|:--------|:----------|:---|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    namedItem.load('type');
    return ctx.sync().then(function() {
            console.log(namedItem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
