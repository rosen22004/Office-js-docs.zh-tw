# <a name="worksheetprotection-object-javascript-api-for-excel"></a>WorksheetProtection 物件 (適用於 Excel 的 JavaScript API)

表示 Sheet 物件的保護。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|protected|bool|表示工作表是否受到保護。唯讀。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|工作表保護選項。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[protect(options:WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|void|保護工作表。如果工作表已經受到保護，則會失敗。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[unprotect()](#unprotect)|void|取消保護工作表。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


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

### <a name="protectoptions-worksheetprotectionoptions"></a>protect(options:WorksheetProtectionOptions)
保護工作表。如果工作表已經受到保護，則會失敗。

#### <a name="syntax"></a>語法
```js
worksheetProtectionObject.protect(options);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|:---|
|選項|WorksheetProtectionOptions|選擇性。工作表保護選項。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    var range = sheet.getRange("A1:B3").format.protection.locked = false;
    sheet.protection.protect({allowInsertRows:true});
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

```
### <a name="unprotect"></a>unprotect()
取消保護工作表。

#### <a name="syntax"></a>語法
```js
worksheetProtectionObject.unprotect();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void
