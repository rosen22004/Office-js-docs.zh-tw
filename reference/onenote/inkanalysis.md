# InkAnalysis 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_   


表示一組指定的筆跡線條的筆跡分析資料。

## 屬性

| 屬性	     | 類型	   |說明|意見反應|
|:---------------|:--------|:----------|:-------|
|id|string|取得 InkAnalysis 物件的識別碼。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysis-id)|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
| 關聯性 | 類型	   |說明| 意見反應|
|:---------------|:--------|:----------|:-------|
|頁面|[頁面](page.md)|取得父頁面物件。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysis-page)|
|段落|[InkAnalysisParagraphCollection](inkanalysisparagraphcollection.md)|取得這個頁面中的筆跡分析段落。 唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysis-paragraphs)|

## 方法

| 方法           | 傳回類型    |說明| 意見反應|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysis-load)|

## 方法詳細資料


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
### 屬性存取範例

**段落**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load ink paragraphs.
    page.load('inkAnalysisOrNull/paragraphs');
    
    return ctx.sync()
        .then(function() {
            console.log(page.inkAnalysisOrNull.paragraphs.items.length);
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```