# OfficeExtension.Error 物件 (適用於 Word 的 JavaScript API)

表示當您使用 Word JavaScript API 時發生的錯誤。

_適用版本：Word 2016、Word for iPad、Word for Mac_

## 屬性
| 屬性	     | 類型	   |說明
|:---------------|:--------|:----------|
|code|string|取得指出錯誤類型的值。 值可以是 "AccessDenied"、"GeneralException"、"ActivityLimitReached"、"InvalidArgument"、"ItemNotFound" 或 "NotImplemented"。 <!-- Values come from OfficeExtension.Error and Word.ErrorCodes. -->|
|debugInfo|string|取得指出當錯誤發生時，會發生什麼事的值。這個值只適用於在開發/偵錯期間。  |
|訊息 |string| 取得對應於錯誤程式碼之當地人們可以讀取的字串。|
|name |string| 取得永遠是 "OfficeExtension.Error" 的值。 |
|traceMessages |string[]| 取得對應於 context.trace(); 設定之檢測訊息的值陣列 |

_請參閱屬性存取[範例。](#範例)_

## 方法

| 方法           | 傳回類型    |說明|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|以下列格式傳回錯誤碼和訊息值："{0}: {1}", code, message。|

## 方法詳細資料

### toString()
以下列格式傳回錯誤碼和訊息值："{0}: {1}", code, message。

#### 語法
```js
error.toString()
```

#### 參數
無

#### 會傳回
字串

#### 範例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    // This will cause an OfficeExtension.Error.
    body.insertText(0);

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync();
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Error code and message: ' + error.toString());
    }
});

```

## 屬性存取範例

### 追蹤訊息檢測

下列範例會顯示如何檢測命令批次檔以判斷發生錯誤的位置。第一批次順利插入文件中的前兩個段落，沒有發生任何錯誤。第二批次順利插入第三個和第四個段落，但在插入第五段落的呼叫時失敗。在批次中失敗的命令之後的所有其他命令都不會執行，包括新增第五個追蹤訊息的命令。在此情況下，可以判斷錯誤是發生於插入第四個段落之後，而在新增第五個追蹤訊息之前。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    // Start a batch of commands.
    body.insertParagraph('1st paragraph', Word.InsertLocation.end);
    // Queue a command for instrumenting this part of the batch.
    context.trace('1st paragraph successful');

    body.insertParagraph('2nd paragraph', Word.InsertLocation.end);
    context.trace('2nd paragraph successful');

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Queue a commmand to insert the paragraph at the end of the document body.
        // Start a new batch of commands.
        body.insertParagraph('3rd paragraph', Word.InsertLocation.end);
        context.trace('3rd paragraph successful');

        body.insertParagraph('4th paragraph', Word.InsertLocation.end);
        context.trace('4th paragraph successful');

        // This command will cause an error. The trace messages in the queue up to
        // this point will be available via Error.traceMessages.
        body.insertParagraph(0, '5th paragraph', Word.InsertLocation.end);
        // Queue a command for instrumenting this part of the batch.
        // This trace message will not be set on Error.traceMessages.
        context.trace('5th paragraph successful');
    }).then(context.sync);
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Trace messages: ' + error.traceMessages);
    }
});

// Output: "Trace messages: 3rd paragraph successful,4th paragraph successful"

```
