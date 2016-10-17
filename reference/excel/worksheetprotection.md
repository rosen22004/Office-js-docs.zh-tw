# <a name="worksheetprotection-object-(javascript-api-for-excel)"></a>WorksheetProtection 物件 (適用於 Excel 的 JavaScript API)

_適用於：Excel 2016、Excel Online、Excel for iOS、Office 2016_

代表 Sheet 物件的保護。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|保護|bool|表示工作表是否受到保護。唯讀。|

## <a name="relationships"></a>關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|選項|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|工作表保護選項。唯讀。|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以工作表的保護詳細資料填入 Proxy 物件。|
|[protect(options:WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoption)|void|保護工作表。如果工作表已經受到保護，則會擲回。|
|[unprotect()](#unprotect)|void|取消保護工作表|

## <a name="method-details"></a>方法詳細資料


### <a name="load(param:-object)"></a>load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### <a name="syntax"></a>語法
```js
object.load(param);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
這個範例會載入使用中工作表的保護詳細資料。
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection status: " + worksheet.protection.protected);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="protect(options:-worksheetprotectionoptions)"></a>protect(options:WorksheetProtectionOptions)
使用選用保護原則保護工作表。如果工作表已經受到保護，則會擲回例外狀況。 

當指定選項時，可以切換啟用或停用個別原則。如果未指定原則，則預設為啟用。 

#### <a name="syntax"></a>語法
```js
worksheetProtectionObject.protect(options);
```

#### <a name="parameters"></a>參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
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
### <a name="unprotect()"></a>unprotect()
取消保護工作表。 

#### <a name="syntax"></a>語法
```js
worksheetProtectionObject.unprotect();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");  
    sheet.protection.unprotect();
    return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```