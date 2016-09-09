
# 在 Outlook 中撰寫約會或郵件時，在本文中插入資料

您可以使用非同步方法 ([Body.getAsync](../../reference/outlook/Body.md)、[Body.getTypeAsync](../../reference/outlook/Body.md)、[Body.prependAsync](../../reference/outlook/Body.md)、[Body.setAsync](../../reference/outlook/Body.md) 和 [Body.setSelectedDataAsync](../../reference/outlook/Body.md)) 來取得本文類型，並在使用者正在撰寫的約會或郵件項目的本文中插入資料。這些非同步方法僅供撰寫增益集使用。若要使用這些方法，請確定您已正確設定增益集資訊清單，以便 Outlook 在撰寫表單中啟動您的增益集，如[建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)中所述。

在 Outlook 中，使用者可以文字、HTML 或 RTF 格式建立郵件，並且可以 HTML 格式建立約會。在插入之前，您應該一律先呼叫 **getTypeAsync**以確認支援的項目格式，因為您可能需要採取其他步驟。**getTypeAsync** 傳回的值取決於原始的項目格式，以及裝置作業系統的支援和以 HTML 格式 (1) 編輯主機。然後據以 (2) 設定 _prependAsync_ 或 **setSelectedDataAsync** 的 **coercionType** 參數以插入資料，如下列表格所示。如果您沒有指定引數，**prependAsync** 和 **setSelectedDataAsync** 會假設要插入的資料是文字格式。



|**要插入的資料**|**getTypeAsync 所傳回的項目格式**|**使用此 coercionType**|
|:-----|:-----|:-----|
|文字|文字 (1)|文字|
|HTML|文字 (1)|文字 (2)|
|文字|HTML|Text/HTML|
|HTML|HTML |HTML|

1.  在平板電腦和智慧型手機上，若作業系統或主機不支援最初在 HTML 中以 HTML 格式建立的編輯項目，則 **getTypeAsync** 會傳回 **Office.MailboxEnums.BodyType.Text**。

2.  如果您要插入的資料是 HTML 且 **getTypeAsync** 傳回該項目的文字類型，將您的資料重新組織為文字，並將其與 **Office.MailboxEnums.BodyType.Text** 插入為 _coercionType_。如果您只要以文字的強制型轉型別插入 HTML 資料，主機會將 HTML 標記顯示為文字。如果您嘗試以 **Office.MailboxEnums.BodyType.Html** 插入 HTML 資料做為 _coercionType_，您會收到錯誤。

除了 _coercionType_ 之外，如同適用於 Office 的 JavaScript API 中大部分的非同步方法，**getTypeAsync**、**prependAsync** 及 **setSelectedDataAsync** 會取得其他選擇性的輸入參數。如需有關指定這些選擇性輸入參數的詳細資訊，請參閱 [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)中的[將選擇性參數傳遞至非同步方法](../../docs/develop/asynchronous-programming-in-office-add-ins.md)。


## 若要在目前的游標位置插入資料


本章節會顯示使用 **getTypeAsync** 的程式碼範例來驗證所撰寫的項目主體類型，然後使用 **setSelectedDataAsync** 在目前的游標位置中插入資料。

您可以將回撥方法及選擇性的輸入參數傳遞至 **getTypeAsync**，並取得 _asyncResult_ 輸出參數中的任何狀態及結果。如果方法成功，您可以取得 [AsyncResult.value](../../reference/shared/asyncresult.status.md) 屬性 (即「文字」或「html」) 中的項目本文的類型。

您必須將資料字串做為輸入參數傳遞至 **setSelectedDataAsync**。依項目本文的類型而定，您可以據以將這個資料字串指定為文字或 HTML 格式。如上所述，您可以選擇性地指定插入 _coercionType_ 參數中的資料類型。此外，您可以提供回撥方法及任何其參數做為選擇性的輸入參數。

如果使用者尚未將游標放置在項目本文中，**setSelectedDataAsync** 會在本文的頂端插入資料。如果使用者已在項目本文中選取文字，**setSelectedDataAsync** 會將選取的文字取代為您指定的資料。請注意，如果使用者要在撰寫項目時同時變更游標位置，則 **setSelectedDataAsync** 可能會失敗。您一次可以插入的字元最大數目為 1,000,000 個字元。

這個程式碼範例假設增益集資訊清單中啟動約會或郵件撰寫表單中的增益集的規則，如下所示。




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 若要在項目本文的開頭插入資料


或者，您可以使用 **prependAsync** 在項目本文的開頭插入資料，並忽略目前的游標位置。除了插入的點之外，**prependAsync** 和 **setSelectedDataAsync** 的行為類似︰


- 如果您要在郵件本文前面加上 HTML 資料，需要先檢查郵件本文的類型以避免以文字格式加上 HTML 資料至郵件。
    
- 提供如下做為對 **prependAsync** 的輸入參數：文字或 HTML 格式及選擇性地要插入資料的格式的資料字串，回撥方法及任何其參數。
    
- 您一次可以在前面加上的字元最大數目為 1,000,000 個字元。
    
下列 JavaScript 程式碼是範例增益集的一部分，在約會和郵件的撰寫格式中啟動。如果項目為約會或 HTML 郵件，範例會呼叫 **getTypeAsync** 以檢查項目本文的類型、將 HTML 資料插入項目本文的頂端，否則以文字格式插入資料。




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 其他資源



- [在 Outlook 中取得並設定撰寫格式的項目資料](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [取得並設定讀取或撰寫格式的 Outlook 項目資料](../outlook/item-data.md)
    
- [建立撰寫格式的 Outlook 增益集](../outlook/compose-scenario.md)
    
- [Office 增益集中的非同步程式設計](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [在 Outlook 中撰寫約會或郵件時，取得、設定或新增收件者。](../outlook/get-set-or-add-recipients.md)
    
- [在 Outlook 中撰寫約會或郵件時，取得或設定主旨](../outlook/get-or-set-the-subject.md)
    
- [在 Outlook 中撰寫約會時，取得或設定位置](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [在 Outlook 中撰寫約會時，取得或設定時間](../outlook/get-or-set-the-time-of-an-appointment.md)
    
