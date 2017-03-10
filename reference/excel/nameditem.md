# <a name="nameditem-object-javascript-api-for-excel"></a>NamedItem 物件 (適用於 Excel 的 JavaScript API)

代表某個儲存格範圍或值的已定義名稱。名稱可以是原始命名物件 (如下列類型所示)、範圍物件、範圍的參照。此物件可用來取得與名稱相關聯的範圍物件。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|註解|string|表示與此名稱相關聯的註解。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|名稱|string|物件的名稱。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|範圍|string|表示名稱是否限於活頁簿或或特定工作表。唯讀。可能的值為：Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|類型|string|表示名稱的公式所傳回的值的類型。唯讀。可能的值為：String、Integer、Double、Boolean、Range。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|物件|代表名稱公式所計算的值。如為具名的範圍，則會傳回範圍位址。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|指定物件是否可見。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱屬性存取[範例。](#property-access-examples)_

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|工作表|[工作表](worksheet.md)|傳回具名項目限於其中的工作表。如果項目改為限於活頁簿，則擲回錯誤。唯讀。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetOrNullObject|[工作表](worksheet.md)|傳回具名項目限於其中的工作表。如果項目改為限於活頁簿，則傳回 null 物件。唯讀。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|刪除指定的名稱。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|傳回與名稱相關的 Range 物件。如果具名項目的類型不是範圍，則擲回錯誤。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeOrNullObject()](#getrangeornullobject)|[範圍](range.md)|傳回與名稱相關的 Range 物件。如果具名項目的類型不是範圍，則傳回 null 物件。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="delete"></a>delete()
刪除指定的名稱。

#### <a name="syntax"></a>語法
```js
namedItemObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

### <a name="getrange"></a>getRange()
傳回與名稱相關的 Range 物件。如果具名項目的類型不是範圍，則擲回錯誤。

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


### <a name="getrangeornullobject"></a>getRangeOrNullObject()
傳回與名稱相關的 Range 物件。如果具名項目的類型不是範圍，則傳回 null 物件。

#### <a name="syntax"></a>語法
```js
namedItemObject.getRangeOrNullObject();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[範圍](range.md)
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
