# Word 增益集程式設計概觀

_適用版本：Word 2016、Word for iPad、Word for Mac_

Word 2016 推出可與 Word 物件搭配使用的新物件模型。此物件模型是 Office.js 所提供之現有物件模型的增補部分，可用來建立適用於 Word 的增益集。您可透過 Web 應用程式裝載的 JavaScript 存取這個物件模型。

## 資訊清單

這個新的 Word 增益集 JavaScript API 使用與 Office 2013 增益集模型相同的資訊清單格式。資訊清單描述裝載增益集的位置、其顯示方式、權限，以及其他資訊。深入了解如何自訂[增益集資訊清單](https://msdn.microsoft.com/en-us/library/office/fp161044.aspx)。 

有數個可用來發佈 Word 增益集資訊清單的選項。閱讀如何[發佈 Office 增益集](https://msdn.microsoft.com/EN-US/library/office/fp123515.aspx)至網路共用、應用程式目錄或 Office 市集。

## 了解適用於 Word 的 JavaScript API

適用於 Word 的 JavaScript API 由 Office.js 載入。它提供一組 JavaScript proxy 物件，用來佇列與 Word 文件內容互動的一組指令。這些命令以批次執行。批次的結果是針對 Word 文件所採取的動作，例如插入內容，以及將 Word 物件與 JavaScript proxy 物件同步處理。 

### 執行增益集

現在來看看執行增益集時需要什麼項目。所有增益集都應該有 Office.initialize 事件處理常式。如需增益集初始化的詳細資訊，請閱讀[了解 API](https://msdn.microsoft.com/EN-US/library/fp160953.aspx)。  

Word 增益集執行時，會在 Word.run() 方法中傳遞函數。Run 方法中傳遞的函數必須具有 context 引數。此 [context 物件](word-add-ins-javascript-reference/requestcontext.md)不同於您從 Office 物件取得的內容物件，雖然其用途相同，都是與 Word 執行階段環境互動。Context 物件可提供對 Word JavaScript 物件模型的存取。現在來看看基本 Word 增益集的註解和程式碼：

**範例 1. 初始化和執行 Word 增益集**

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time the page is loaded.
        Office.initialize = function (reason) {
            
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason 
                // argument to determine how the add-in was loaded.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
        })
    })();
```

範例 1. 示範建立 Word 增益集所需的基本程式碼。它會初始化 Office.js，並包含與 Word 文件互動的 run 方法。

### Proxy 物件

Word JavaScript 物件模型與 Word 中的物件鬆散結合。Word JavaScript 物件是 Word 文件中實際物件的 proxy 物件。除非同步處理文件狀態，否則對 proxy 物件採取的所有動作都不會在 Word 中實現，而 Word 文件的狀態也不會在 proxy 物件中實現。執行 context.sync() 時就會同步處理文件狀態。Sync() 方法基本上會對每一個 proxy 物件執行佇列中的命令組。範例 2 示範建立 proxy body 物件與佇列命令，以載入 proxy body 物件上的 text 屬性，接著將 Word 文件的內文與 body proxy 物件同步處理。 

**範例 2. 將文件內文與 body proxy 物件同步處理。**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values. 
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

### 命令佇列

Word proxy 物件具有可以存取和更新物件模型的方法。系統會依照批次中方法的佇列順序，循序執行這些方法。呼叫 context.sync() 之前，應先形成一個命令批次。使用此 context 之所有物件中佇列的命令都會加以執行。  

在範例 3 中，我們將示範命令佇列的運作方式。呼叫 context.sync() 時，發生的第一件事是在 Word 中執行[命令以載入](Word%20Add-ins%20JavaScript%20Reference/loadoption.md)內文文字。接著，執行將文字插入 Word 內文的命令。隨後將結果傳回 body proxy 物件。Word JavaScript 中 body.text 屬性的值，將是文字插入 Word 文件<u>之前</u>的 Word 文件內文的值。 

**範例 3. 執行命令批次。**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text in the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

## 歡迎您提供意見

我們很重視您的意見。 

* 查看文件，並在此存放庫中直接[送出問題](https://github.com/OfficeDev/office-js-docs/issues)，即可告知我們您找到的任何問題。
* 請告訴我們您的程式設計經驗、您希望未來版本提供哪些功能、程式碼範例等等。使用[這個網站](http://officespdev.uservoice.com/)可輸入您的建議和想法。


## 其他資源

* [Word 增益集](word-add-ins.md)
* [Word 增益集 JavaScript 參考](word-add-ins-javascript-reference.md)
* [Office 增益集](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office 增益集入門](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;GitHub 上的 Word 增益集&lt;/a&gt;
* [Word 的程式碼片段總管](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)

