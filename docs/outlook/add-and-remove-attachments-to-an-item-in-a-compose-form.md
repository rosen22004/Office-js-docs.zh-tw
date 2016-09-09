
# 在 Outlook 中新增及移除撰寫格式項目的附件

您可以使用 [addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 和 [addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法，分別將檔案和 Outlook 項目附加到使用者撰寫的項目。兩者都是非同步方法，這表示不需等到新增附件動作完成，便可以繼續執行。根據原始位置以及要新增的附件大小，新增附件非同步呼叫可能需要一段時間才能完成。如果有仰賴動作完成的工作，您應該在回撥方法中執行這些工作。這個回撥方法是選擇性的，並且是在附件上載完成時叫用。回撥方法需要 [AsyncResult](http://dev.outlook.com/reference/add-ins/simple-types.md) 物件當做輸出參數，從新增附件動作提供任何狀態、錯誤和傳回值。如果回撥需要任何額外的參數，您可以在選擇性的 _options.aysncContext_ 參數中指定這些字元。_options.asyncContext_ 可以是回撥方法所預期的任何類型。

例如，您可以將 _options.asyncContext_ 定義為包含一或多個機碼值組的 JSON 物件，利用 ':' 字元來分隔機碼和值，以及利用 ',' 來分隔不同的機碼值組。您可以在 [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)的 Office 增益集平台中找到[將選擇性參數傳遞至非同步方法](../../docs/develop/asynchronous-programming-in-office-add-ins.md)的更多範例。下列範例示範如何使用 **asyncContext** 參數來將 2 個引數傳遞至回撥方法︰




```js
{ asyncContext: { var1: 1, var2: 2} }
```

您可以使用 **AsyncResult** 物件的 **status** 和 **error** 屬性來檢查非同步方法呼叫中的回撥方法為成功或錯誤。如果附加成功完成，您可以使用 **AsyncResult.value** 屬性，以取得附件識別碼。附件識別碼是整數，您後續可以用來移除附件。


 >**附註：**最佳作法是，唯有當相同增益集已在相同的工作階段中新增該附件時，您才應該使用附件識別碼來移除附件。在 Outlook Web App 和 OWA for Devices 中，附件識別碼只有在相同工作階段內才會有效。當使用者關閉增益集時，工作階段會結束，或如果使用者開始在內嵌表單進行撰寫，接下來會跳出內嵌表單，以便在個別視窗中繼續。


## 附加檔案

您可以透過使用 **addFileAttachmentAsync** 方法並指定檔案的 URI，在撰寫表單中將檔案附加到郵件或約會。如果檔案受到保護，則您可以併入適當的識別或驗證權杖，作為 URI 查詢字串參數。Exchange 會對 URI 進行呼叫以取得附件，而保護檔案的 Web 服務將需要使用權杖作為驗證方法。

下列 JavaScript 範例為撰寫增益集，其會從 Web 伺服器附加檔案 picture.png 至要撰寫的郵件或約會中。回撥方法會取得 **asyncResult** 作為參數、檢查附加狀態，並取得附件識別碼 (如果附加成功)。




```js
var mailbox;
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

Office.initialize = function () {
    mailbox = Office.context.mailbox;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID. 
        // You can optionally pass any object that you would  
        // access in the callback method as an argument to  
        // the asyncContext parameter.
        mailbox.item.addFileAttachmentAsync(
            attachmentURI,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 附加 Outlook 項目

您可以藉由指定項目的 Exchange Web 服務 (EWS) 識別碼，並使用 **addItemAttachmentAsync** 方法，將 Outlook 項目 (例如，電子郵件、行事曆或連絡人項目) 附加到撰寫表單的郵件或約會中。您也可以使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法，並存取 EWS 作業 [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)，來取得使用者信箱中電子郵件、行事曆、連絡人或工作項目的 EWS 識別碼。[Item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.md) 屬性也會在讀取表單中提供現有項目的 EWS 識別碼。

下列的 JavaScript 函式 `addItemAttachment` 會延伸上述的第一個範例，並加入項目，作為要撰寫的電子郵件或約會的附件。此函式會取得要附加之項目的 EWS 識別碼作為引數。如果附加成功，它會取得附件識別碼來作進一步的處理，包括在相同的工作階段中移除該附件。




```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(ID) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.addItemAttachmentAsync(
        ID,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```


 >**附註：**您可以使用撰寫增益集，在 Outlook Web App 或裝置用 OWA 中附加週期性約會的執行個體。不過，在支援的 Outlook 豐富型用戶端中，嘗試附加執行個體會造成附加週期性序列 (主要約會)。


## 移除附件


您可以藉由指定相對應的附件識別碼，並使用 [removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法，從撰寫表單的郵件或約會項目中移除檔案或項目附件。您僅應該移除相同增益集在相同工作階段中加入的附件。您應該要確定附件識別碼對應到有效的附件，否則方法會傳回錯誤。類似於 **addFileAttachmentAsync** 和 **addItemAttachmentAsync** 方法，**removeAttachmentAsync** 是非同步方法。您應該提供回撥方法，以使用 **AsyncResult** 輸出參數物件來檢查狀態和任何錯誤。您也可以利用選擇性 **asyncContext** 參數 (也就是機碼值組的 JSON 物件) 傳遞任何其他參數至回撥方法。

下列的 JavaScript 函式 `removeAttachment` 會繼續延伸上述的範例，並從要撰寫的電子郵件或約會中移除指定的附件。此函式會取得要移除之附件的識別碼作為引數。在 **addFileAttachmentAsync** 或 **addItemAttachmentAsync** 方法呼叫成功之後，您可以取得的附件的識別碼，並針對後續的 **removeAttachmentAsync** 方法呼叫儲存它。




```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be 
// removed. 
function removeAttachment(ID) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.removeAttachmentAsync(
        ID,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```


## 加入和移除附件的提示


如果您的撰寫增益集會加入和移除附件，請建構程式碼，以便將有效的附件識別碼傳遞給 remove-attachment 呼叫，並處理 **AsyncResult.error** 傳回 **InvalidAttachmentId** 時的情況。根據附件的位置及大小，附加檔案或項目可能需要一些時間才能完成。下列範例包含對 **addFileAttachmentAsync**、`write` 和 **removeAttachmentAsync** 的呼叫。您可能會認為呼叫會依序一個接著一個執行。


```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

// Gets the current time in minutes, seconds and milliseconds.
function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);
            }
            write ('(3): ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
    'attachmentID is: ' + attachmentID);

Office.context.mailbox.item.removeAttachmentAsync(
        attachmentID,      
        { asyncContext: null },
       function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(5): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {           
                write('(6): ' + minutesSecondsMilliSeconds() + ' ' + 
                    ID of removed attachment: ' + asyncResult.value);
            }
        });


```

即使 **addFileAttachmentAsync** 是在 **removeAttachmentAsync** 前開始，因為 **addFileAttachmentAsync** 為非同步，`write` 和 **removeAttachmentAsync** 呼叫可以在 **addFileAttachmentAsync** 完成之前開始。發生這種情況時，`attachmentID` 會保持 **未定義**，而您會收到 **removeAttachmentAsync** 呼叫的錯誤，如以下的輸出︰




```
 (4): 46:18:245 attachmentID is: undefined
Error executing code: Sys.ArgumentException: Sys.ArgumentException: Value does not fall within the expected range. Parameter name: attachmentId
 (2): 46:18:255 ID of added attachment: 0
 (3): 46:18:262 Finishing addFileAttachmentAsync callback method.
```

避免此情況的一個方式是在呼叫 **removeAttachmentAsync** 之前，先檢查已定義 `attachmentID`。另一個方式是從 **addFileAttachmentAsync** 回撥方法內初始化 **removeAttachmentAsync** 呼叫，如下列範例所示︰




```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1) ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2) ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);

                // Move the write and removeAttachmentAsync calls here 
                // inside the addFileAttachmentAsync callback, after the 
                // attaching has succeeded.
                write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'attachmentID is: ' + attachmentID);

                Office.context.mailbox.item.removeAttachmentAsync(
                    attachmentID,
                    { asyncContext: null },
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            write('(5) ' + minutesSecondsMilliSeconds() + ' ' + 
                                asyncResult.error.message);
                        }
                        else {
                            write('(6) ' + minutesSecondsMilliSeconds() + ' ' + 
                                'ID of removed attachment: ' + attachmentID);
                        }
                    });
            }

            write('(3) ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

以下是輸出的一個範例：




```
(2) 49:25:775 ID of added attachment: 1
(4) 49:25:782 attachmentID is: 1
(3) 49:25:783 Finishing addFileAttachmentAsync callback method.
(6) 49:25:789 ID of removed attachment: 1
```

請注意，**removeAttachmentAsync** 的回撥會為 **addFileAttachmentAsync** 回撥內形成巢狀。因為 **addFileAttachmentAsync** 和 **removeAttachmentAsync** 為非同步，**addFileAttachmentAsync** 的回撥中的最後一行可以在 **removeAttachmentAsync** 的回撥完成之前便執行。


## 其他資源



- [建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)
    
- [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    


