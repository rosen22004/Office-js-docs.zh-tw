# PageContentCollection 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


代表頁面的內容，是 PageContent 物件的集合。

## 屬性

| 屬性	     | 類型	   |說明|意見反應|
|:---------------|:--------|:----------|:-------|
|Count|int|傳回集合中的頁面內容數目。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-count)|
|項目|[PageContent[]](pagecontent.md)|PageContent 物件的集合。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-items)|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明| 意見反應|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number 或 string)](#getitemindex-number-或-string)|[PageContent](pagecontent.md)|藉由識別碼或藉由其集合中的索引，來取得 PageContent 物件。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[PageContent](pagecontent.md)|根據頁面內容在集合中的位置，取得頁面內容。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-load)|

## 方法詳細資料


### getItem(index: number 或 string)
藉由識別碼或藉由其集合中的索引，來取得 PageContent 物件。唯讀。

#### 語法
```js
pageContentCollectionObject.getItem(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|Index|number 或 string|PageContent 物件的識別碼，或其在集合中的索引位置。|

#### 傳回
[PageContent](pagecontent.md)

### getItemAt(index: number)
根據頁面內容在集合中的位置，取得頁面內容。

#### 語法
```js
pageContentCollectionObject.getItemAt(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[PageContent](pagecontent.md)

#### 範例
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;
    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("The first page content item is of type: " + firstPageContent.type);
            return context.sync();
        });
})
.catch(function(error) {
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
### 屬性存取範例

**項目**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Queue a command to load the type of each pageContent.
    pageContents.load("type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            $.each(pageContents.items, function(index, pageContent) {
                console.log("PageContent type: " + pageContent.type);
            });
        });
})                
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**周遊大綱**
```js
OneNote.run(function (context) {
   var page = context.application.getActivePage();
   var pageContents = page.contents;
   pageContents.load('type');
   var outlines = [];
   return context.sync()
       .then(function () {    
              $.each(pageContents.items, function (index, pageContent) {
                     console.log(pageContent.type);
                     if (pageContent.type === 'Outline') {
                           outlines.push(pageContent);
                     }
              });
              $.each(outlines, function (index, outline) {
                     outline.load("id,paragraphs,paragraphs/type");
              });
              return context.sync();
       })
       .then(function () {
              $.each(outlines, function (index, outline) {
                     console.log("An outline was found with id : " + outline.id);
              });
              return Promise.resolve(outlines);
       });
});
```

