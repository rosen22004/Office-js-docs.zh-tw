
# 從 Outlook 項目擷取實體字串

本文說明如何建立**顯示實體** Outlook 增益集，其可擷取選取的 Outlook 項目的主旨及本文中的受支援已知實體的字串執行個體。這個項目可以是約會、電子郵件或會議邀請、回應或取消。支援的實體包括︰

- 地址：包含至少有街道號碼、街道名稱、城市和郵遞區號等元素之子集的美國郵寄地址。
    
- 聯絡人：在其他實體的內容中 (例如地址或公司名稱) 包含人員的連絡人資訊。
    
- 電子郵件地址：SMTP 電子郵件地址。
    
- 會議建議︰會議建議，例如事件的參考。請注意，只有郵件 (而非約會) 支援解壓縮會議建議。
    
- 電話號碼：北美地區的電話號碼。
    
- 工作建議：工作的建議，通常是在可採取行動的片語中表示。
    
- URL。
    
這些實體大部分仰賴根據大量資料的機器學習為基礎的自然語言辨識。這個辨識不具決定性，且有時取決於 Outlook 項目中的內容。每當使用者選擇檢視一個約會、電子郵件或會議邀請、回應或取消時，Outlook 會啟動實體增益集。在初始化期間，範例實體增益集會從目前的項目讀取支援實體的所有執行個體。 

增益集為使用者提供可選擇實體類型的按鈕。當使用者選取實體時，增益集會在增益集窗格中顯示已選取實體的執行個體。下列章節列出 XML 資訊清單，以及實體增益集的 HTML 和 JavaScript 檔案，並反白顯示支援個別實體擷取的程式碼。

## XML 資訊清單


實體增益集有兩個由邏輯 OR 運算加入的啟用規則。 


```xml
<!-- Activate the add-in if the current item in Outlook is an email or appointment item. -->
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
</Rule>
```

這些規則會指定當目前在讀取窗格或讀取檢查程式中選取的項目為約會或郵件時 (包含電子郵件或會議要求、回應或取消) 時，Outlook 應該啟動這個增益集。

以下為實體增益集的資訊清單：它會使用 Office 增益集資訊清單的結構描述 1.1 版。




```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
xsi:type="MailApp">
  <Id>6880A140-1C4F-11E1-BDDB-0800200C9A68</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>EN-US</DefaultLocale>
  <DisplayName DefaultValue="Display entities"/>
  <Description DefaultValue=
     "Display known entities on the selected item."/>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify the web -->
        <!-- server where the HTML file is hosted. -->
        <SourceLocation DefaultValue=
          "http://webserver/default_entities/default_entities.html"/>
        <RequestedHeight>350</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <!-- Activate the add-in if the current item in Outlook is -->
  <!-- an email or appointment item. -->
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemIs" ItemType="Appointment"/>
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>
</OfficeApp>
```


## HTML 實作


實體增益集的 HTML 檔案會指定供使用者選取每一種類型實體的按鈕，以及清除執行個體的所顯示實體的另一個按鈕。它包含 JavaScript 檔案 (default_entities.js)，其會在 [JavaScript 實作](#javascript-實作)的下一節中說明。JavaScript 檔案包括每個按鈕的事件處理常式。

請注意，所有的 Outlook 增益集都必須包含 office.js。隨後的 HTML 檔案包含 CDN 上的 1.1 版 office.js 。 

```html
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" >
    <title>standard_item_properties</title>
    <link rel="stylesheet" type="text/css" media="all" href="default_entities.css" />
    <script type="text/javascript" src="MicrosoftAjax.js"></script>
    <!-- Use the CDN reference to Office.js. -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    <script type="text/javascript"  src="default_entities.js"></script>
</head>

<body>
    <div id="container">
        <div id="button">
        <input type="button" value="clear" 
            onclick="myClearEntitiesBox();">
        <input type="button" value="Get Addresses" 
            onclick="myGetAddresses();">
        <input type="button" value="Get Contact Information" 
            onclick="myGetContacts();">
        <input type="button" value="Get Email Addresses" 
            onclick="myGetEmailAddresses();">
        <input type="button" value="Get Meeting Suggestions" 
            onclick="myGetMeetingSuggestions();">
        <input type="button" value="Get Phone Numbers" 
            onclick="myGetPhoneNumbers();">
        <input type="button" value="Get Task Suggestions" 
            onclick="myGetTaskSuggestions();">
        <input type="button" value="Get URLs" 
            onclick="myGetUrls();">
        </div>
        <div id="entities_box"></div>
    </div>
</body>
</html>
```


## 樣式表


實體增益集使用選擇性的 CSS 檔案 (default_entities.css) 來指定輸出的版面配置。以下是 CSS 檔案的清單。


```css
*
{
    color: #FFFFFF;
    margin: 0px;
    padding: 0px;
    font-family: Arial, Sans-serif;
}
html 
{
    scrollbar-base-color: #FFFFFF;
    scrollbar-arrow-color: #ABABAB; 
    scrollbar-lightshadow-color: #ABABAB; 
    scrollbar-highlight-color: #ABABAB; 
    scrollbar-darkshadow-color: #FFFFFF; 
    scrollbar-track-color: #FFFFFF;
}
body
{
    background: #4E9258;
}
input
{
    color: #000000;
    padding: 5px;
}
span
{
    color: #FFFF00;
}
div#container
{
    height: 100%;
    padding: 2px;
    overflow: auto;
}
div#container td
{
    border-bottom: 1px solid #CCCCCC;
}
td.property-name
{
    padding: 0px 5px 0px 0px;
    border-right: 1px solid #CCCCCC;
}
div#meeting_suggestions
{
    border-top: 1px solid #CCCCCC;
}
```


## JavaScript 實作


其餘章節會說明這個範例 (default_entities.js 檔) 如何從使用者正在檢視的郵件或約會的主旨及本文擷取已知實體。 


## 在初始化時擷取實體


在 [Office.initialize](../../reference/shared/office.initialize.md) 事件時，實體增益集會目前項目的呼叫 [getEntities](../../reference/outlook/Office.context.mailbox.item.md) 方法。**getEntities** 方法會傳回全域變數 `_MyEntities` 支援實體的執行個體陣列。以下是相關的 JavaScript 程式碼。


```js
// Global variables
var _Item;
var _MyEntities;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    MyEntities = _Item.getEntities();
    
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}

```


## 擷取位址


當使用者按一下 [取得位址]**** 按鈕時，如果擷取任何位址，`myGetAddresses` 事件處理常式會從 `_MyEntities` 物件的 [addresses](../../reference/outlook/simple-types.md) 屬性取得位址陣列。 每一個擷取的地址會在陣列中儲存為字串。 `myGetAddresses` 在 .mdText` 中形成本機的 HTML 字串以顯示解壓縮地址的清單。 以下是相關的 JavaScript 程式碼。


```js
// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## 擷取連絡人資訊


當使用者按一下 [取得連絡人資訊]**** 按鈕時，`myGetContacts` 事件處理常式會取得一個連絡人的陣列，以及 (若擷取任何項目) 從 `_MyEntities` 物件的 [contacts](../../reference/outlook/simple-types.md) 屬性取得其資訊。 每個截取的連絡人會在陣列中儲存為 [Contact](../../reference/outlook/simple-types.md) 物件。 `myGetContacts` 取得關於每個連絡人的進一步資料。 請注意，內容會決定 Outlook 是否可以從項目 - 電子郵件結尾的簽章擷取連絡人，或至少部分的下列資訊必須存在連絡人的鄰近位置︰


- [Contact.personName](../../reference/outlook/simple-types.md) 屬性中代表連絡人名稱的字串。
    
- [Contact.businessName](../../reference/outlook/simple-types.md) 屬性中代表與連絡人相關聯的公司名稱的字串。
    
- [Contact.phoneNumbers](../../reference/outlook/simple-types.md) 屬性中代表與連絡人相關聯的電話號碼的陣列。每一個電話號碼由 [PhoneNumber](../../reference/outlook/simple-types.md) 物件所代表。
    
- 針對電話號碼陣列中的每一個 **PhoneNumber** 成員，[PhoneNumber.phoneString](../../reference/outlook/simple-types.md) 屬性中代表電話號碼的字串。
    
- [Contact.urls](../../reference/outlook/simple-types.md) 屬性中與連絡人相關聯的 URL 陣列。每個 URL 會以陣列成員中的字串方式表示。
    
- [Contact.emailAddresses](../../reference/outlook/simple-types.md) 屬性中與連絡人相關聯的電子郵件地址的陣列。每個電子郵件地址會以陣列成員中的字串方式表示。
    
- [Contact.addresses](../../reference/outlook/simple-types.md) 屬性中與連絡人相關聯的郵寄地址的陣列。每個郵寄地址會以陣列成員中的字串方式表示。
    
 `myGetContacts` 會在 `htmlText` 中形成本機 HTML 字串來顯示每個連絡人的資料。以下是相關的 JavaScript 程式碼。




```js
// Gets instances of the Contact entity on the item.
function myGetContacts()
{
    var htmlText = "";

    // Gets an array of contacts and their information.
    var contactsArray = _MyEntities.contacts;
    for (var i = 0; i < contactsArray.length; i++)
    {
        // Gets the name of the person. The name is a string.
        htmlText += "Name : <span>" + contactsArray[i].personName +
            "</span><br/>";

        // Gets the company name associated with the contact.
        htmlText += "Business : <span>" + 
        contactsArray[i].businessName + "</span><br/>";

        // Gets an array of phone numbers associated with the 
        // contact. Each phone number is represented by a 
        // PhoneNumber object.
        var phoneNumbersArray = contactsArray[i].phoneNumbers;
        for (var j = 0; j < phoneNumbersArray.length; j++)
        {
            htmlText += "PhoneString : <span>" + 
                phoneNumbersArray[j].phoneString + "</span><br/>";
            htmlText += "OriginalPhoneString : <span>" + 
                phoneNumbersArray[j].originalPhoneString +
                "</span><br/>";
        }

        // Gets the URLs associated with the contact.
        var urlsArray = contactsArray[i].urls;
        for (var j = 0; j < urlsArray.length; j++)
        {
            htmlText += "Url : <span>" + urlsArray[j] + 
                "</span><br/>";
        }

        // Gets the email addresses of the contact.
        var emailAddressesArray = contactsArray[i].emailAddresses;
        for (var j = 0; j < emailAddressesArray.length; j++)
        {
           htmlText += "E-mail Address : <span>" + 
               emailAddressesArray[j] + "</span><br/>";
        }

        // Gets postal addresses of the contact.
        var addressesArray = contactsArray[i].addresses;
        for (var j = 0; j < addressesArray.length; j++)
        {
          htmlText += "Address : <span>" + addressesArray[j] + 
              "</span><br/>";
        }

        htmlText += "<hr/>";
        }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## 截取電子郵件地址


當使用者按一下 [取得電子郵件地址]**** 按鈕時，如果擷取任何位址，`myGetEmailAddresses` 事件處理常式會從 `_MyEntities` 物件的 [emailAddresses](../../reference/outlook/simple-types.md) 屬性取得 SMTP 電子郵件地址陣列。 每一個擷取的電子郵件地址會在陣列中儲存為字串。 `myGetEmailAddresses` 在 `htmlText` 中形成本機的 HTML 字串以顯示解壓縮的電子郵件地址清單。 以下是相關的 JavaScript 程式碼。


```js
// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses() {
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++) {
        htmlText += "E-mail Address : <span>" + emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## 擷取會議建議


當使用者按一下 [取得會議建議]**** 按鈕時，如果擷取任何位址，`myGetMeetingSuggestions` 事件處理常式會從 `_MyEntities` 物件的 [meetingSuggestions](../../reference/outlook/simple-types.md) 屬性取得會議建議的陣列。


 >**附註** 只有郵件 (而非約會) 支援 **MeetingSuggestion** 實體類型。

每個已擷取的會議建議會在陣列中儲存為 [MeetingSuggestion](../../reference/outlook/simple-types.md) 物件。`myGetMeetingSuggestions` 會取得關於每個會議建議的進一步資料：


- 已識別為 [MeetingSuggestion.meetingString](../../reference/outlook/simple-types.md) 屬性中的會議建議的字串。
    
- [MeetingSuggestion.attendees](../../reference/outlook/simple-types.md) 屬性中的會議出席者的陣列。每個出席者由 [EmailUser](../../reference/outlook/simple-types.md) 物件所代表。
    
- 針對每個出席者，[EmailUser.displayName](../../reference/outlook/simple-types.md) 屬性中的名稱。
    
- 針對每個出席者，[EmailUser.emailAddress](../../reference/outlook/simple-types.md) 屬性中的 SMTP 地址。
    
- [MeetingSuggestion.location](../../reference/outlook/simple-types.md) 屬性中代表會議建議位置的字串。
    
- [MeetingSuggestion.subject](../../reference/outlook/simple-types.md) 屬性中代表會議建議主旨的字串。
    
- [MeetingSuggestion.start](../../reference/outlook/simple-types.md) 屬性中代表會議建議開始時間的字串。
    
- [MeetingSuggestion.end](../../reference/outlook/simple-types.md) 屬性中代表會議建議結束時間的字串。
    
 `myGetMeetingSuggestions` 會在 `htmlText` 中形成本機 HTML 字串來顯示每個會議建議的資料。以下是相關的 JavaScript 程式碼。




```js
// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions() {
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++) {
        // Gets the string that was identified as a meeting suggestion.
        htmlText += "MeetingString : <span>" + meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion. Each attendee is represented 
        // by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++) {
            htmlText += "Attendee : ( ";

            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}
```


## 擷取電話號碼


當使用者按一下 [取得電話號碼]**** 按鈕時，如果擷取任何項目，`myGetPhoneNumbers` 事件處理常式會從 `_MyEntities` 物件的 [phoneNumbers](../../reference/outlook/simple-types.md) 屬性取得電話號碼的陣列。 每個截取的電話號碼會在陣列中儲存為 [PhoneNumber](../../reference/outlook/simple-types.md) 物件。 `myGetPhoneNumbers` 取得關於每個電話號碼的進一步資料。


- [PhoneNumber.type](../../reference/outlook/simple-types.md) 屬性中代表電話號碼類型的字串 (例如住家電話號碼)。
    
- [PhoneNumber.phoneString](../../reference/outlook/simple-types.md) 屬性中代表實際電話號碼的字串。
    
- [PhoneNumber.originalPhoneString](../../reference/outlook/simple-types.md) 屬性中原先識別為電話號碼的字串。
    
 `myGetPhoneNumbers` 會在 `htmlText` 中形成本機 HTML 字串來顯示每個電話號碼的資料。以下是相關的 JavaScript 程式碼。




```js
// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
           "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
            phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## 擷取工作建議


當使用者按一下 [取得工作建議]**** 按鈕時，如果擷取任何位址，`myGetTaskSuggestions` 事件處理常式會從 `_MyEntities` 物件的 [taskSuggestions](../../reference/outlook/simple-types.md) 屬性取得工作建議的陣列。 每個已擷取的工作建議會在陣列中儲存為 [TaskSuggestion](../../reference/outlook/simple-types.md) 物件。 `myGetTaskSuggestions` 會取得關於每個工作建議的進一步資料：


- [TaskSuggestion.taskString](../../reference/outlook/simple-types.md) 屬性中原先識別為工作建議的字串。
    
- The array of task assignees from the [TaskSuggestion.assignees](../../reference/outlook/simple-types.md) 屬性中工作受託人的陣列。每個受託人由 [EmailUser](../../reference/outlook/simple-types.md) 物件所代表。
    
- 針對每個受託人，[EmailUser.displayName](../../reference/outlook/simple-types.md) 屬性中的名稱。
    
- 針對每個受託人，[EmailUser.emailAddress](../../reference/outlook/simple-types.md) 屬性中的 SMTP 地址。
    
 `myGetTaskSuggestions` 會在 `htmlText` 中形成本機 HTML 字串來顯示每個工作建議的資料。以下是相關的 JavaScript 程式碼。




```js
// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from 
    // the current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
           tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
               "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## 擷取 URL


當使用者按一下 [取得 URL]**** 按鈕時，如果擷取任何項目，`myGetUrls` 事件處理常式會從 `_MyEntities` 物件的 [urls](../../reference/outlook/simple-types.md) 屬性取得 URL 的陣列。 每一個擷取的 URL 會在陣列中儲存為字串。 `myGetUrls` 在 `htmlText` 中形成本機的 HTML 字串以顯示解壓縮的 URL 清單。


```js
// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## 清除顯示的實體字串


最後，實體增益集會指定清除任何顯示字串的 `myClearEntitiesBox` 事件處理常式。以下是相關的程式碼。


```js
// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}
```


## JavaScript 清單


以下是 JavaScript 實作的完整清單。


```js
// Global variables
var _Item;
var _MyEntities;

// Initializes the add-in.
Office.initialize = function () {
    var _mailbox = Office.context.mailbox;
    // Obtains the current item.
    _Item = _mailbox.item;
    // Reads all instances of supported entities from the subject 
    // and body of the current item.
    _MyEntities = _Item.getEntities();

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    });
}


// Clears the div with id="entities_box".
function myClearEntitiesBox()
{
    document.getElementById("entities_box").innerHTML = "";
}

// Gets instances of the Address entity on the item.
function myGetAddresses()
{
    var htmlText = "";

    // Gets an array of postal addresses. Each address is a string.
    var addressesArray = _MyEntities.addresses;
    for (var i = 0; i < addressesArray.length; i++)
    {
        htmlText += "Address : <span>" + addressesArray[i] + 
            "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the EmailAddress entity on the item.
function myGetEmailAddresses()
{
    var htmlText = "";

    // Gets an array of email addresses. Each email address is a 
    // string.
    var emailAddressesArray = _MyEntities.emailAddresses;
    for (var i = 0; i < emailAddressesArray.length; i++)
    {
        htmlText += "E-mail Address : <span>" + 
            emailAddressesArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the MeetingSuggestion entity on the 
// message item.
function myGetMeetingSuggestions()
{
    var htmlText = "";

    // Gets an array of MeetingSuggestion objects, each array 
    // element containing an instance of a meeting suggestion 
    // entity from the current item.
    var meetingsArray = _MyEntities.meetingSuggestions;

    // Iterates through each instance of a meeting suggestion.
    for (var i = 0; i < meetingsArray.length; i++)
    {
        // Gets the string that was identified as a meeting 
        // suggestion.
        htmlText += "MeetingString : <span>" + 
            meetingsArray[i].meetingString + "</span><br/>";

        // Gets an array of attendees for that instance of a 
        // meeting suggestion.
        // Each attendee is represented by an EmailUser object.
        var attendeesArray = meetingsArray[i].attendees;
        for (var j = 0; j < attendeesArray.length; j++)
        {
            htmlText += "Attendee : ( ";
            // Gets the displayName property of the attendee.
            htmlText += "displayName = <span>" + attendeesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each attendee.
            // This is the SMTP address of the attendee.
            htmlText += "emailAddress = <span>" + attendeesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        // Gets the location of the meeting suggestion.
        htmlText += "Location : <span>" + 
            meetingsArray[i].location + "</span><br/>";

        // Gets the subject of the meeting suggestion.
        htmlText += "Subject : <span>" + 
            meetingsArray[i].subject + "</span><br/>";

        // Gets the start time of the meeting suggestion.
        htmlText += "Start time : <span>" + 
           meetingsArray[i].start + "</span><br/>";

        // Gets the end time of the meeting suggestion.
        htmlText += "End time : <span>" + 
            meetingsArray[i].end + "</span><br/>";

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}


// Gets instances of the phone number entity on the item.
function myGetPhoneNumbers()
{
    var htmlText = "";

    // Gets an array of phone numbers. 
    // Each phone number is a PhoneNumber object.
    var phoneNumbersArray = _MyEntities.phoneNumbers;
    for (var i = 0; i < phoneNumbersArray.length; i++)
    {
        htmlText += "Phone Number : ( ";
        // Gets the type of phone number, for example, home, office.
        htmlText += "type = <span>" + phoneNumbersArray[i].type + 
            "</span> , ";

        // Gets the actual phone number represented by a string.
        htmlText += "phone string = <span>" + 
            phoneNumbersArray[i].phoneString + "</span> , ";

        // Gets the original text that was identified in the item 
        // as a phone number. 
        htmlText += "original phone string = <span>" + 
           phoneNumbersArray[i].originalPhoneString + "</span>";

        htmlText += " )<br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the task suggestion entity on the item.
function myGetTaskSuggestions()
{
    var htmlText = "";

    // Gets an array of TaskSuggestion objects, each array element 
    // containing an instance of a task suggestion entity from the 
    // current item.
    var tasksArray = _MyEntities.taskSuggestions;

    // Iterates through each instance of a task suggestion.
    for (var i = 0; i < tasksArray.length; i++)
    {
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
            tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
            htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
                "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
                "</span>";

            htmlText += " )<br/>";
        }

        htmlText += "<hr/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

// Gets instances of the URL entity on the item.
function myGetUrls()
{
    var htmlText = "";

    // Gets an array of URLs. Each URL is a string.
    var urlArray = _MyEntities.urls;
    for (var i = 0; i < urlArray.length; i++)
    {
        htmlText += "Url : <span>" + urlArray[i] + "</span><br/>";
    }

    document.getElementById("entities_box").innerHTML = htmlText;
}

```


## 其他資源



- [建立讀取格式的 Outlook 增益集](../outlook/read-scenario.md)
    
- [使 Outlook 項目中的字串與已知的實體相符](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [item.getEntities 方法](../../reference/outlook/Office.context.mailbox.item.md)
    
