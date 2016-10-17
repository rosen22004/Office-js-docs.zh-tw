
# <a name="get,-set,-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>在 Outlook 中撰寫約會或郵件時，取得、設定或新增收件者。


適用於 Office 的 JavaScript API 分別提供非同步方法 ([Recipients.getAsync](../../reference/outlook/Recipients.md)、[Recipients.setAsync](../../reference/outlook/Recipients.md) 或 [Recipients.addAysnc](../../reference/outlook/Recipients.md)) 給約會或郵件的撰寫表單中的取得、設定或加入收件者。這些非同步方法僅供撰寫增益集使用。若要使用這些方法，請確定您已正確設定適用於 Outlook 的增益集資訊清單以啟動撰寫表單中的增益集，如[建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)中所述。

部分代表約會或郵件中收件者的屬性在讀取表單和撰寫表單中可提供讀取權限使用。這些屬性包括 [optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) 和 [requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md) (適用於約會)，及 [cc](../../reference/outlook/Office.context.mailbox.item.md) 和 [to](../../reference/outlook/Office.context.mailbox.item.md) (適用於郵件)。在讀取表單中，您可以直接從父物件存取屬性，例如︰




```js
item.cc
```

但在撰寫表單中，因為使用者及您的增益集可能會同時插入或變更收件者，您必須使用非同步方法 **getAsync** 來取得這些屬性，如下列範例中所示︰




```js
item.cc.getAsync
```

這些屬性僅在撰寫表單中，而非在閱讀表單中可供撰寫權限使用。

如同適用於 Office 的 JavaScript APIe 中大部分的非同步方法，**getAsync**、**setAsync** 和 **addAsync** 接受選擇性輸入參數。如需有關指定這些選擇性輸入參數的詳細資訊，請參閱 [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)中的[將選擇性參數傳遞至非同步方法](../../docs/develop/asynchronous-programming-in-office-add-ins.md)。


## <a name="to-get-recipients"></a>若要取得收件者


本章節會顯示取得正在撰寫的約會或郵件的收件者，及顯示收件者的電子郵件地址的程式碼範例。程式碼範例假設增益集資訊清單中啟動約會或郵件撰寫表單中的增益集的規則，如下所示。 


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

在適用於 Office 的 JavaScript API 中，因為代表約會的收件者的屬性 (**optionalAttendees** 和 **requiredAttendees**) 與郵件的屬性 ([bcc](../../reference/outlook/Office.context.mailbox.item.md)、**cc** 和 **to**) 不同，您應該先使用 [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) 屬性來判斷正在撰寫的項目是否為約會或郵件。在撰寫模式中，這些約會及郵件的屬性全部為 [Recipients](../../reference/outlook/Recipients.md) 物件，因此您可以套用非同步方法 **Recipients.getAsync** 來取得對應的收件者。 

若要使用 **getAsync**，提供回撥方法以檢查由非同步 **getAsync** 呼叫傳回的狀態、結果及任何錯誤。您可以使用可選的 _asyncContext_ 參數提供任何引數給回撥方法。回撥方法會傳回 _asyncResult_ 輸出參數。您可以使用 **AsyncResult** 參數物件的 **status** 和 [error](../../reference/outlook/simple-types.md) 屬性來檢查非同步呼叫的狀態及任何錯誤訊息，以及 **value** 屬性以取得實際的收件者。收件者會以 [EmailAddressDetails](../../reference/outlook/simple-types.md) 物件的陣列做為代表。

請注意，因為 **getAsync** 方法是非同步，若有依成功取得收件者的後續動作，當非同步呼叫成功完成時，您應該僅在對應的回撥方法中組織您的程式碼來啟動這些動作。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="to-set-recipients"></a>若要設定收件者


本章節會顯示設定使用者正在由使用者撰寫的約會或郵件的收件者的程式碼範例。設定收件者會覆寫任何現有的收件者。類似於先前取得撰寫表格中收件者的範例，此範例假設會在約會及郵件的撰寫表單中啟動增益集。本範例首先驗證撰寫的項目為約會還是郵件，因此在代表約會或郵件收件者的適當屬性上套用非同步方法 **Recipients.setAsync**。

當呼叫 **setAsync** 時，提供陣列做為 _recipients_ 參數的輸入引數，使用下列其中一種格式：


- SMTP 地址的字串陣列。
    
- 字典的陣列，每個都包含顯示名稱和電子郵件地址，如下列程式碼範例中所示。
    
- **EmailAddressDetails** 物件的陣列，類似於 **getAsync** 方法所傳回的陣列。
    
您可以選擇性地提供回撥方法做為輸入引數至 **setAsync** 方法，以確定依成功設定收件者的任何程式碼僅在這種情況發生時才會執行。您也可以使用可選的 _asyncContext_ 參數提供任何引數給回撥方法。若您使用回撥方法，可以存取 _asyncResult_ 輸出參數，並使用 **AsyncResult** 參數物件的 **status** 和 **error** 屬性來檢查非同步呼叫的狀態及任何錯誤訊息。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## <a name="to-add-recipients"></a>若要新增收件者


如果您不想覆寫約會或郵件中的任何現有收件者，您可以使用 **Recipients.addAsync** 非同步方法來附加收件者，而非使用 **Recipients.setAsync**。**addAsync** 運作方式類似 **setAsync** 因為其需要 _recipients_ 輸入引數。您可以使用 asyncContext 參數，選擇性地提供回撥方法及回撥的任何引數。然後您可以使用回撥方法的 **asyncResult** 輸出參數來檢查非同步 _addAsync_ 回呼的狀態、結果及任何錯誤。下列範例會檢查正在撰寫的項目是否為約會，並且將兩個必要的出席者附加至約會。


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## <a name="additional-resources"></a>其他資源



- [在 Outlook 中取得並設定撰寫格式的項目資料](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [取得並設定讀取或撰寫格式的 Outlook 項目資料](../outlook/item-data.md)
    
- [建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)
    
- [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [在 Outlook 中撰寫約會或郵件時，取得或設定主旨](../outlook/get-or-set-the-subject.md)
    
- [在 Outlook 中撰寫約會或郵件時，在本文中插入資料](../outlook/insert-data-in-the-body.md)
    
- [在 Outlook 中撰寫約會時，取得或設定位置](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [在 Outlook 中撰寫約會時，取得或設定時間](../outlook/get-or-set-the-time-of-an-appointment.md)
    
