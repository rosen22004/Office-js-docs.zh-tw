# <a name="table-object-javascript-api-for-excel"></a>Table 物件 (適用於 Excel 的 JavaScript API)

代表 Excel 表格。

## <a name="properties"></a>屬性

| 屬性	       | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|highlightFirstColumn|bool|指出第一個資料行是否包含特殊格式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|highlightLastColumn|bool|指出最後一個資料行是否包含特殊格式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|id|int|傳回可唯一識別特定活頁簿中表格的值。即使將表格重新命名，識別碼的值仍保持不變。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|表格的名稱。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedColumns|bool|表示資料行是否顯示帶狀格式，其中奇數的資料行會以不同於偶數資料行的方式反白顯示，讓閱讀表格更方便。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedRows|bool|表示資料列是否顯示帶狀格式，其中奇數的資料列會以不同於偶數資料列的方式反白顯示，讓閱讀表格更方便。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showFilterButton|bool|表示篩選按鈕是否在各個資料行標頭上方可見。只有在表格包含標題列時允許設定這個選項。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showHeaders|bool|指出是否顯示標題列。可以設定此值，以顯示或移除標題列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTotals|bool|指出是否顯示合計列。可以設定此值，以顯示或移除合計列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|字串|代表表格樣式的常數值。可能的值為：TableStyleLight1 到 TableStyleLight21、TableStyleMedium1 到 TableStyleMedium28、TableStyleStyleDark1 到 TableStyleStyleDark11。也可以指定在活頁簿中呈現自訂的使用者定義樣式。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_請參閱關於效能 (與具有[公式](#setting-formulas)_的表格相關) 的重要注意事項



## <a name="relationships"></a>關聯性
| 關聯性 | 類型	    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|columns|[TableColumnCollection](tablecolumncollection.md)|傳回表格中所有資料行的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rows|[TableRowCollection](tablerowcollection.md)|傳回表格中所有列的集合。唯讀。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[TableSort](tablesort.md)|代表表格的排序。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|包含目前表格的工作表。唯讀。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 傳回類型    |描述| 需求集合|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|清除目前在表格上套用的所有篩選器。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToRange()](#converttorange)|[Range](range.md)|將表格轉換成一般儲存格範圍。所有的資料會保留。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|刪除表格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|取得與表格的資料主體相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|取得與表格的標題列相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|取得與整個表格相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|取得與表格的合計列相關的範圍物件。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapplyFilters()](#reapplyfilters)|void|重新套用目前在表格上的所有篩選器。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法詳細資料


### <a name="clearfilters"></a>clearFilters()
清除目前在表格上套用的所有篩選器。

#### <a name="syntax"></a>語法
```js
tableObject.clearFilters();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

### <a name="converttorange"></a>convertToRange()
將表格轉換成一般儲存格範圍。所有的資料會保留。

#### <a name="syntax"></a>語法
```js
tableObject.convertToRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="delete"></a>delete()
刪除表格。

#### <a name="syntax"></a>語法
```js
tableObject.delete();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange"></a>getDataBodyRange()
取得與表格的資料主體相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableObject.getDataBodyRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableDataRange = table.getDataBodyRange();
    tableDataRange.load('address')
    return ctx.sync().then(function() {
            console.log(tableDataRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getheaderrowrange"></a>getHeaderRowRange()
取得與表格的標題列相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableObject.getHeaderRowRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('address');
    return ctx.sync().then(function() {
        console.log(tableHeaderRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrange"></a>getRange()
取得與整個表格相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableObject.getRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load('address');    
    return ctx.sync().then(function() {
            console.log(tableRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange"></a>getTotalRowRange()
取得與表格的合計列相關的範圍物件。

#### <a name="syntax"></a>語法
```js
tableObject.getTotalRowRange();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
[Range](range.md)

#### <a name="examples"></a>範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableTotalsRange = table.getTotalRowRange();
    tableTotalsRange.load('address');    
    return ctx.sync().then(function() {
            console.log(tableTotalsRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="reapplyfilters"></a>reapplyFilters()
重新套用目前在表格上的所有篩選器。

#### <a name="syntax"></a>語法
```js
tableObject.reapplyFilters();
```

#### <a name="parameters"></a>參數
無

#### <a name="returns"></a>傳回
void
### <a name="property-access-examples"></a>屬性存取範例

依名稱取得表格。 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('index')
    return ctx.sync().then(function() {
            console.log(table.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

依索引取得表格。

```js
Excel.run(function (ctx) { 
    var index = 0;
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('id')
    return ctx.sync().then(function() {
            console.log(table.id);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

設定表格樣式。 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.name = 'Table1-Renamed';
    table.showTotals = false;
    table.style = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.style);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="setting-formulas"></a>設定公式

#### <a name="common-pitfalls-when-setting-formulas-in-excel-from-add-ins"></a>從增益集在 Excel 中設定公式時的常見錯誤

作者：Zlatko Michailov  
Microsoft Corp.


這份文件指出 Excel 增益集開發人員可能會遇到的三個陷阱及其解決方法。請務必了解這些案例，尤其是在正常情況下它們不會造成增益集失敗。增益集使用於小範圍時會看似很正常，但是它們可能會在目標範圍隨時間不斷增長時，以線性方式下降。

在__表格__資料行上設定公式時，前兩項會顯示資訊清單，更精確來說，就是具有公式的資料行以及具有總資料列的資料行。

##### <a name="setting-formulas-in-calculated-table-columns"></a>在計算的表格資料行中設定公式

此[文件](https://support.office.com/en-us/article/Use-calculated-columns-in-an-Excel-table-873FBAC6-7110-4300-8F6F-AAFA2EA11CE8)提供計算資料行的概觀。

主要功能是在步驟 4 中︰

> 當您按 Enter 時，公式會自動填滿資料行的所有儲存格 — 以上以及您用來輸入公式的儲存格底下。公式也適用於每個資料列，但因為它是結構化參照，Excel 會知道哪一個資料列是哪一個。

這表示每個單一的公式更新可能會收到 N 次，其中 N 是在表格中的資料列數目。

使用者可能不會注意到處理 1000 個資料列的表格時，會發生長時間延遲，但與包含 10000 這些資料列的表格互動時，可能會遭受影響。

幸好，Excel 的自動資料行計算非常聰明，所以您可能不會注意到上述問題。若要取得自動重新計算的資料行，它必須是空白或者是完全自動計算。如果您將值插入任何儲存格內，來中斷資料行的「純正性」(不是公式) 時，Excel 將不會嘗試自動重新計算。此外，如果您嘗試設定 Excel 已在該資料行所設的公式，重新計算不會執行任何作業。

在範例中，我們假設您正在`=B2+C2`儲存格上`A2`設定公式。如果資料行是空的，Excel 會計算此資料行的所有儲存格並_調整資料列索引_。接著，當您將移至下一列中，然後將公式 `=B3+C3` 設定在 `A3` 上，將不重新計算任何資料行，因為此公式已經自動設定在全部資料行上。

不過，如果您希望您的資料行呈現資料列索引的函數，例如 `=i * i`，_i_ 是資料列索引，且不只會在每次更新時重新計算整個資料行，您的資料行最後也會顯示相同的 (上一個) 公式。

##### <a name="setting-formulas-on-a-table-with-a-totals-row"></a>在具有「總計」資料列的表格設定公式

在啟用總計資料列的表格上設定公式，可能有時會造成效能問題。請務必注意，即便預設總計資料列也無法重現問題，像是在最左邊的儲存格具有靜態值，以及在最右邊的儲存格具有 `Count`，且在 `None` 之間具有所有儲存格。 

還是有簡單的解決方法，請設定所有公式，然後將總計資料列新增至表格上，建議上述問題的一般因應措施模式，是在設定公式時使用一般範圍，然後將該範圍轉換成表格。

這是能更新資料範圍的泛型函數，並能在目標範圍上建立表格。 

```js
function createAndPopulateTable(context, worksheetName, rangeAddress, hasHeaderRow, headerValues, bodyFormulas, tableCustomizer) {
    var worksheet = context.workbook.worksheets.getItem(worksheetName);

    // Calculate table-, body-, and header- ranges
    var tableRange = worksheet.getRange(rangeAddress);
    var bodyRange = tableRange;
    if (hasHeaderRow) {
        bodyRange = tableRange.getResizedRange(-1, 0).getOffsetRange(1, 0);
        if (headerValues) {
            // Set header values
            var headerRange = tableRange.getRow(0);
            headerRange.values = headerValues;
        }
    }
    
    // Set body formulas
    bodyRange.formulas = bodyFormulas;

    return context.sync()
        .then(function() {
            // Create the table
            var table = context.workbook.tables.add(tableRange, hasHeaderRow);

            // Invoke the caller's customizer
            if (tableCustomizer) {
                tableCustomizer(table);
            }

            return context.sync();
        });
}
```

上述函數能在線上的[公用位置](https://gist.github.com/zlatko-michailov/2b0418c986d9da6ee0bdf7aa346d3a4f)使用。

它可以像這樣使用︰
```js
    return Excel.run(function(context) {
        return createAndPopulateTable(context, "Sheet1", "B3:E6", true, [['Alpha', 'Beta', 'Gamma', 'Delta']], 
                    [ ['=1+1', null, null, '=B4'], 
                      ['=2+2', null, null, '=B5'],
                      ['=3+3', null, null, '=B6'] ],
                    function (table) {
                        table.style = 'TableStyleLight1';
                        table.showTotals = true;
                    });
    });
```

可以在 Excel 桌面用戶端 (依預設為開啟) 停用自動資料行計算，但其實它在 Excel Online 中永遠為開啟。因此，身為增益集開發人員，針對大部分的增益集的使用者，您應該假設它為開啟。


##### <a name="getting-a-range-object"></a>取得範圍物件

這個問題是特別針對 JavaScript API 實作。

為了能在插入與刪除資料列/資料行期間正確追蹤範圍，繫結會在每次要求 `Range` 物件時於內部中建立。稍後，當更新儲存格時，必須通知所有相關的繫結自行更新。

因此，下列程式碼 (第 8 行)，從一般的程式設計觀點來說不具威脅性，會逐層向上提報複雜性：
```js
    Excel.run(function(context) {
        var n = 10000;
        var worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load();

        var arr = [];
        for (var i = 2; i <= n + 1; i++) {
            var range = worksheet.getRange("C3:C" + (n + 1)); /* <-- PROBLEM! */
            arr.push(["=A" + i + " + B" + i]);
        }
        range.formulas = arr; 
        return context.sync();
    });
```

因應措施藉由將相關行排除在迴圈外，來避免將不必要的 get 放到相同的 `Range` 物件：
```js
    Excel.run(function(context) {
        var n = 10000;
        var worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load();

        var arr = [];
        var range = worksheet.getRange("C3:C" + (n + 1)); /* <-- OK */
        for (var i = 2; i <= n + 1; i++) {
            arr.push(["=A" + i + " + B" + i]);
        }
        range.formulas = arr; 
        return context.sync();
    });
```
