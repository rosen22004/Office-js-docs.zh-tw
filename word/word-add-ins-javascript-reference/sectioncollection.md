# SectionCollection 物件 (適用於 Word 的 JavaScript API)

包含文件的 [section](section.md) 物件集合。

_適用版本：Word 2016、Word for iPad、Word for Mac_

## 屬性
| 屬性	   | 類型	|說明
|:---------------|:--------|:----------|
|項目|[Section[]](section.md)|Section 物件的集合。唯讀。|

_請參閱屬性存取[範例。](#property-access-examples)_

## 關聯性
無


## 方法

| 方法		   | 傳回類型	|說明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。|

## 方法詳細資料

### load(param: object)
以參數中指定的屬性和物件值填滿 JavaScript 層中建立的 Proxy 物件。

#### 語法
```js
object.load(param);
```

#### 參數
| 參數	   | 類型	|說明|
|:---------------|:--------|:----------|
|param|object|選用。接受參數與關聯性名稱，做為分隔字串或陣列。或者提供 [loadOption](loadoption.md) 物件。|

#### 傳回
void

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary header of the first section. 
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");
        
        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
        });                    
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

## 支援詳細資料

在執行階段檢查使用[需求集](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)以確認您的應用程式受到 Word 主應用程式版本的支援。如需有關 Office 主應用程式及伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)。 