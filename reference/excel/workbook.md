# Workbook 物件 (適用於 Excel 的 JavaScript API)

活頁簿是最上層物件，其中包含相關的活頁簿物件，例如工作表、表格、範圍等等。

## 屬性

無

## 關聯性
| 關聯性 | 類型	   |說明|
|:---------------|:--------|:----------|
|應用程式|[應用程式](application.md)|代表包含此活頁簿的 Excel 應用程式執行個體。唯讀。|
|bindings|[BindingCollection](bindingcollection.md)|代表屬於活頁簿一部份的繫結集合。唯讀。|
|函數|[函數](functions.md)|代表包含此活頁簿的 Excel 應用程式執行個體。唯讀。|
|names|[NamedItemCollection](nameditemcollection.md)|代表活頁簿限定範圍具名項目 (具名的範圍和常數) 的集合。唯讀。|
|tables|[TableCollection](tablecollection.md)|代表與活頁簿關聯的表格集合。唯讀。|
|worksheets|[WorksheetCollection](worksheetcollection.md)|代表與活頁簿關聯的工作表集合。唯讀。|

## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[範圍](range.md)|從活頁簿取得目前選取的範圍。|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料


### getSelectedRange()
從活頁簿取得目前選取的範圍。

#### 語法
```js
workbookObject.getSelectedRange();
```

#### 參數
無

#### 傳回
[範圍](range.md)

#### 範例

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|param|物件|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void
