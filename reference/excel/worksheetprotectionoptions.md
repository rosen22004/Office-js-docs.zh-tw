# <a name="worksheetprotectionoptions-object-(javascript-api-for-excel)"></a>WorksheetProtectionOptions 物件 (適用於 Excel 的 JavaScript API)

_適用於：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示工作表保護中的選項。

## <a name="properties"></a>屬性

| 屬性	     | 類型	   |描述
|:---------------|:--------|:----------|
|allowAutoFilter|bool|代表工作表保護選項，允許使用自動篩選功能。|
|allowDeleteColumns|bool|代表工作表保護選項，允許刪除資料行。|
|allowDeleteRows|bool|代表工作表保護選項，允許刪除資料列。|
|allowFormatCells|bool|代表工作表保護選項，允許格式化儲存格。|
|allowFormatColumns|bool|代表工作表保護選項，允許格式化資料行。|
|allowFormatRows|bool|代表工作表保護選項，允許格式化資料列。|
|allowInsertColumns|bool|代表工作表保護選項，允許插入資料行。|
|allowInsertHyperlinks|bool|代表工作表保護選項，允許插入超連結。|
|allowInsertRows|bool|代表工作表保護選項，允許插入資料列。|
|allowPivotTables|bool|代表工作表保護選項，允許使用樞紐分析表功能。|
|allowSort|bool|代表工作表保護選項，允許使用排序功能。|

_請參閱屬性存取[範例。](#examples)_

## <a name="relationships"></a>關聯性
無


## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

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
這個範例會載入使用中工作表的保護選項。
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection options: " + worksheet.protection.options);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
