# PageCollection 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


代表頁面的集合。

## 屬性

| 屬性	     | 類型	   |說明|意見反應|
|:---------------|:--------|:----------|:-------|
|Count|int|傳回集合中的頁面數目。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-count)|
|項目|[Page[]](page.md)|Page 物件的集合。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-items)|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明| 意見反應|
|:---------------|:--------|:----------|:-------|
|[getByTitle(title: string)](#getbytitletitle-string)|[PageCollection](pagecollection.md)|取得具有指定名稱的頁面集合。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getByTitle)|
|[getItem(index: number 或 string)](#getitemindex-number-或-string)|[Page](page.md)|藉由識別碼或藉由其集合中的索引，來取得頁面。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Page](page.md)|根據頁面在集合中的位置，取得頁面。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-load)|

## 方法詳細資料


### getByTitle(title: string)
取得具有指定名稱的頁面集合。

#### 語法
```js
pageCollectionObject.getByTitle(title);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|標題|string|頁面的標題。|

#### 傳回
[PageCollection](pagecollection.md)

#### 範例
```js
OneNote.run(function (context) {

    // Get all the pages in the current section.
    var allPages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.
    allPages.load("id"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Get the sections with the specified name.
            var todoPages = allPages.getByTitle("Todo list");

            // Queue a command to load the section. 
            // For best performance, request specific properties.
            todoPages.load("id,title"); 

            return context.sync()
                .then(function () {

                    // Iterate through the collection or access items individually by index.
                    if (todoPages.items.length > 0) {
                        console.log("Page title: " + todoPages.items[0].title);
                        console.log("Page ID: " + todoPages.items[0].id);
                    }
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

### getItem(index: number 或 string)
藉由識別碼或藉由其集合中的索引，來取得頁面。唯讀。

#### 語法
```js
pageCollectionObject.getItem(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|Index|number 或 string|頁面的識別碼，或頁面在集合中的索引位置。|

#### 傳回
[Page](page.md)

### getItemAt(index: number)
根據頁面在集合中的位置，取得頁面。

#### 語法
```js
pageCollectionObject.getItemAt(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[Page](page.md)

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
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
    
    // Queue a command to load the id and title for each page.            
    pages.load('id,title');
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Display the properties.
            $.each(pages.items, function(index, page) {
                console.log(page.title);
                console.log(page.id);
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

