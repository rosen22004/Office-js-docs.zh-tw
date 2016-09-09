# TableSort 物件 (適用於 Excel 的 JavaScript API)

_適用版本：Excel 2016、Excel Online、Excel for iOS、Office 2016_

管理在 Table 物件上的排序作業。

## 屬性

| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|matchCase|bool|表示大小寫會影響料表的最後排序。唯讀。|
|方法|string|表示最後用於排序資料表的中文字元排序方法。唯讀。可能的值為：拼音、StrokeCount。|

## 關聯性
| 關聯性 | 類型	   |描述|
|:---------------|:--------|:----------|
|欄位|[SortField](sortfield.md)|表示用於最後排序資料表的目前條件。唯讀。|

## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[apply(fields:SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|執行排序作業。|
|[clear()](#clear)|void|清除資料表上目前的排序。雖然這不會修改資料表的順序，它會清除標頭按鈕的狀態。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|
|[reapply()](#reapply)|void|將目前的排序參數重新套用至資料表。|

## 方法詳細資料


### apply(fields:SortField[], matchCase: bool, method: string)
執行排序作業。

#### 語法
```js
tableSortObject.apply(fields, matchCase, method);
```

#### 參數
| 參數	    | 類型	   |描述|
|:---------------|:--------|:----------|
|欄位|SortField[]|用來排序的條件清單。|
|matchCase|bool|選擇性。是否有大小寫影響的字串排序。|
|方法|string|選擇性。適用於中文字元的排序方法。可能的值為：拼音、StrokeCount|

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
清除資料表上目前的排序。雖然這不會修改資料表的順序，它會清除標頭按鈕的狀態。

#### 語法
```js
tableSortObject.clear();
```

#### 參數
無

#### 傳回
void

#### 範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### Syntax
```js
object.load(param);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void

### reapply()
將目前的排序參數重新套用至資料表。

#### 語法
```js
tableSortObject.reapply();
```

#### 參數
無

#### 傳回
void

####範例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.reapply();   
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});