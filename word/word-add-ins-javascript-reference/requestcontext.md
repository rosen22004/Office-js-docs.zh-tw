# RequestContext 物件 (適用於 Word 的 JavaScript API)

RequestContext 物件可協助從 Word 增益集向 Word 應用程式提出要求，因為這兩個應用程式在不同的處理程序中執行。 

_適用版本：Word 2016、Word for iPad、Word for Mac_

## 屬性
無

## 方法

| 方法         | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |以參數中指定的屬性和選項填滿 JavaScript 層中建立的 proxy 物件。|
|[sync()](#sync)  |Promise 物件 |送出要求佇列給 Word，並傳回 promise 物件，此物件可用於鏈結進一步動作。|

## 方法詳細資料

### load(object: object, option: object)
以參數中指定的屬性和選項填滿 JavaScript 層中建立的 proxy 物件。

#### 語法
```js
requestContextObject.load(object, loadOption);
```

#### 參數
| 參數       | 類型    |描述|
|:----------------|:--------|:----------|
|物件|object|選用。指定要載入之物件的名稱。|
|option|[loadOption](loadoption.md)|選用，但這是最佳作法。指定載入選項，例如 select、expand、skip 和 top。 |

#### 傳回
void

##### 範例

下列範例示範如何使用要求內容來載入段落集合上的文字屬性。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
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

### sync() 
送出要求佇列給 Word，並傳回 promise 物件，此物件可用於鏈結進一步動作。

#### 語法
```js
requestContextObject.sync();
```

#### 參數
無

#### 傳回
Promise 物件。

#### 範例

下列範例示範使用兩次同步方法：1) 載入內容控制項集合與每個內容控制項的文字屬性，以及 2) 清除集合中的第一個內容控制項的內容。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
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
