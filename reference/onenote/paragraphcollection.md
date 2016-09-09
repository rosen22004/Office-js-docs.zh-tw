# ParagraphCollection 物件 (適用於 OneNote 的 JavaScript API)

_適用於：OneNote Online_  


代表 Paragraph 物件的集合。

## 屬性

| 屬性	     | 類型	   |說明|意見反應|
|:---------------|:--------|:----------|:-------|
|Count|int|傳回頁面中的段落數目。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-count)|
|項目|[Paragraph[]](paragraph.md)|Paragraph 物件的集合。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-items)|

_請參閱屬性存取[範例。](#範例)_

## 關聯性
無


## 方法

| 方法           | 傳回類型    |說明| 意見反應|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number 或 string)](#getitemindex-number-或-string)|[段落](paragraph.md)|藉由識別碼或藉由其集合中的索引，來取得 Paragraph 物件。唯讀。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[段落](paragraph.md)|根據段落在集合中的位置，取得段落。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|[執行](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-load)|

## 方法詳細資料


### getItem(index: number 或 string)
藉由識別碼或藉由其集合中的索引，來取得 Paragraph 物件。唯讀。

#### 語法
```js
paragraphCollectionObject.getItem(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|Index|number 或 string|Paragraph 物件的識別碼，或其在集合中的索引位置。|

#### 傳回
[段落](paragraph.md)

### getItemAt(index: number)
根據段落在集合中的位置，取得段落。

#### 語法
```js
paragraphCollectionObject.getItemAt(index);
```

#### 參數
| 參數	    | 類型	   |說明|
|:---------------|:--------|:----------|
|index|number|要擷取之物件的索引值。以 0 開始編製索引。|

#### 傳回
[段落](paragraph.md)

#### 範例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page, and then get its Outline's first paragraph.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;

    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the type and richText.text property of this paragraph.
    firstParagraph.load("id,type");


    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            // Write text from paragraph to console
            console.log("First Paragraph found with id : " + firstParagraph.id + " and type " + firstParagraph.type);
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

    // Get the first PageContent on the page, and then get its Outline's first paragraph.
    var pageContent = pageContents.getItem(0);
    var paragraphs = pageContent.outline.paragraphs;
    
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            var firstParagraph = paragraphs.items[0];
            // Write text from first paragraph to console
            console.log("First Paragraph found with id : " + firstParagraph.id + " and type " + firstParagraph.type);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**周遊 richText**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page, and then get its outline's paragraphs.
    var outlinePageContents = [];
    var paragraphs = [];
    var richTextParagraphs = [];
    // Queue a command to load the id and type of each page content in the outline.
    pageContents.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            // Load all page contents of type Outline
            $.each(pageContents.items, function(index, pageContent) {
                if(pageContent.type == 'Outline')
                {
                    pageContent.load('outline,outline/paragraphs,outline/paragraphs/type');
                    outlinePageContents.push(pageContent);
                }
            });
            return context.sync();
        })
        .then(function () {
            // Load all rich text paragraphs across outlines
            $.each(outlinePageContents, function(index, outlinePageContent) {
                var outline = outlinePageContent.outline;
                paragraphs = paragraphs.concat(outline.paragraphs.items);
            });
            $.each(paragraphs, function(index, paragraph) {
                if(paragraph.type == 'RichText')
                {
                    richTextParagraphs.push(paragraph);
                    paragraph.load("id,richText/text");
                }
            });
            return context.sync();
        })
        .then(function () {
            // Display all rich text paragraphs to the console
            $.each(richTextParagraphs, function(index, richTextParagraph) {
                var richText = richTextParagraph.richText;
                console.log("Paragraph found with richtext content : " + richText.text + " and richtext id : " + richText.id);
            });
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

