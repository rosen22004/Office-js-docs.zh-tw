
# <a name="get-the-whole-document-from-an-add-in-for-powerpoint-or-word"></a>從 PowerPoint 或 Word 增益集中，取得整份文件

您可以建立 Office 增益集，按一下即可將 Word 2013 或 PowerPoint 2013 文件傳送或發佈到遠端位置。本文示範如何建置 PowerPoint 2013 的簡單工作窗格增益集，將所有簡報當成一個資料物件，並透過 HTTP 要求將該資料傳送到網頁伺服器。

## <a name="prerequisites-for-creating-an-add-in-for-powerpoint-or-word"></a>建立 PowerPoint 或 Word 增益集的必要條件


本文假設您使用文字編輯器來建立 PowerPoint 或 Word 的工作窗格增益集。若要建立工作窗格增益集，您必須建立下列檔案︰


- 在共用網路資料夾或網頁伺服器上，您需要下列檔案︰
    
      - HTML 檔案 (GetDoc_App.html)，包含使用者介面加上 JavaScript 檔案 (包括 office.js 和主機特定的 .js 檔案) 和階層式樣式表 (CSS) 檔案的連結。
    
  - JavaScript 檔案 (GetDoc_App.js)，以包含增益集的程式設計邏輯。
    
  - 包含增益集的樣式與格式的 CSS 檔案 (Program.css)。
    
- 增益集的 XML 資訊清單檔案 (GetDoc_App.xml)，可在共用網路資料夾或增益集目錄上使用。資訊清單檔必須指向先前所述的 HTML 檔案的位置。
    
您也可以使用 Visual Studio 2015 來建立 PowerPoint 或 Word 的增益集。 


### <a name="core-concepts-to-know-for-creating-a-task-pane-add-in"></a>建立工作窗格增益集需要了解的核心概念

開始為 PowerPoint 或 Word 建立此增益集之前，您應該熟悉建置 Office 增益集及使用 HTTP 要求。本文不討論如何從網頁伺服器的 HTTP 要求中解碼 Base64 編碼的文字。 


## <a name="create-the-manifest-for-the-add-in"></a>建立增益集的資訊清單


PowerPoint 增益集的 XML 資訊清單檔案提供增益集的重要資訊︰有哪些應用程式可以裝載它、HTML 檔案的位置、增益集標題和描述，以及許多其他特性。


- 在文字編輯器中，將下列程式碼新增至資訊清單檔案。
    
```XML
  
<?xml version="1.0" encoding="utf-8" ?> 
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="TaskPaneApp">
    <Id>[Replace_With_Your_GUID]</Id> 
    <Version>1.0</Version> 
    <ProviderName>[Provider Name]</ProviderName> 
    <DefaultLocale>EN-US</DefaultLocale> 
    <DisplayName DefaultValue="Get Doc add-in" /> 
    <Description DefaultValue="My get PowerPoint or Word document add-in." /> 
    <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" /> 
    <Hosts>
      <Host Name="Document" /> 
      <Host Name="Presentation" /> 
    </Hosts>
    <DefaultSettings>
      <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" /> 
    </DefaultSettings>
    <Permissions>ReadWriteDocument</Permissions> 
</OfficeApp>
```

- 將使用 UTF-8 編碼的 GetDoc_App.xml 檔案儲存至網路位置或增益集目錄。
    

## <a name="create-the-user-interface-for-the-add-in"></a>建立增益集的使用者介面


對於增益集的使用者介面，您可以使用 HTML，直接寫入至 GetDoc_App.html 檔案。增益集的程式設計邏輯和功能必須包含在 JavaScript 檔案中 (例如，GetDoc_App.js)。

您可以使用下列程序來建立增益集的簡單使用者介面，其中包含一個標題和一個按鈕。


- 在文字編輯器的新檔案中，加入下列 HTML。
    
```html  
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
        <title>Publish presentation</title>
        <link rel="stylesheet" type="text/css" href="Program.css" />
        <script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="GetDoc_App.js"></script>
    </head>
    <body>
      <form>
        <h1>Publish presentation</h1>
        <br />
        <div><input id='submit' type="button" value="Submit" /></div>
        <br />
        <div><h2>Status</h2> 
            <div id="status"></div>
        </div>
      </form>
    </body>
</html>
```

- 使用 UTF-8 編碼將檔案儲存為 GetDoc_App.xml 至網路位置或網頁伺服器。
    

 >**附註：**確定增益集的 **head** 標記包含的 **script** 標記具有 office.js 檔案的有效連結。 

我們將使用一些 CSS 來取得簡單，但時尚並具有專業外觀的增益集。您可以使用下列 CSS 來定義增益集的樣式。


- 在文字編輯器的新檔案中，加入下列 CSS。
    
```css 
body
{
    font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
}
h1,h2
{
    text-decoration-color:#4ec724;
}
input [type="submit"], input[type="button"] 
{ 
    height:24px; 
    padding-left:1em; 
    padding-right:1em; 
    background-color:white; 
    border:1px solid grey; 
    border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0; 
    cursor:pointer; 
}
```

- 使用 UTF-8 編碼將檔案儲存為 Program.css 至網路位置或 GetDoc_App.html 檔案所在的網頁伺服器。
    

## <a name="add-the-javascript-to-get-the-document"></a>加入 JavaScript 以取得文件


在增益集的程式碼中，[Office.initialize](../../reference/shared/office.initialize.md) 事件的處理常式會將處理常式加入至表單上 [提交]**** 按鈕的 click 事件，並通知使用者增益集已就緒。

下列程式碼範例示範 **Office.initialize** 事件的事件處理常式以及 helper 函式 `updateStatus`，用於寫入至狀態 div。




```js
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

      // After the DOM is loaded, add-in-specific code can run.
      document.getElementById('submit').addEventListener("click",
          function () {
              sendFile();
          });}
      updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div. 
function updateStatus(message) {
    var statusInfo = document.getElementById("status");
    statusInfo.innerHTML += message + "<br/>";
}
```



當您在 UI 中選擇 [提交]**** 按鈕時，增益集會呼叫 `sendFile` 函式，其包含對 [Document.getFileAsync](../../reference/shared/document.getfileasync.md) 方法的呼叫。**getFileAsync** 方法會使用非同步模式，類似於適用於 Office 的 JavaScript API 中的其他方法。它有一個必要參數 _fileType_，和兩個選擇性參數，_options_ 和 _callback_。 

_fileType_ 參數預期 [FileType](../../reference/shared/filetype-enumeration.md) 列舉的三個常數之一：**Office.FileType.Compressed** ("compressed")、**Office.FileType.PDF** ("pdf") 或 **Office.FileType.Text** ("text")。PowerPoint 只支援 **Compressed** 作為引數；Word 支援全部三個。當您傳入 **fileType** 參數的 _Compressed_ 時，**getFileAsync** 方法會在本機電腦上建立檔案的暫存副本，傳回文件作為 PowerPoint 2013 簡報檔案 (*.pptx) 或 Word 2013 文件檔案 (*.docx)。

**getFileAsync** 方法會傳回檔案參考來作為 [File](../../reference/shared/file.md) 物件。**File** 物件公開四個成員︰[size](../../reference/shared/file.size.md) 屬性、[sliceCount](../../reference/shared/file.slicecount.md) 屬性、[getSliceAsync](../../reference/shared/file.getsliceasync.md) 方法和 [closeAsync](../../reference/shared/file.closeasync.md) 方法。**size** 屬性傳回檔案中的位元組數目。**SliceCount** 傳回檔案中的 [Slice](../../reference/shared/document.md) 物件數目 (本文稍後討論)。

下列程式碼使用 **document.getFileAsync()** 方法擷取 PowerPoint 或 Word 文件做為 **File** 物件。然後封裝所產生的 **File** 物件、清空的計數器以及 [sliceCount](../../reference/shared/file.slicecount.md) 到匿名的物件。這個物件為後續傳遞給本機定義的 `getSlice` 函式。 

```js
// Get all the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {

    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status == Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size +
                    " bytes");

                getSlice(state);
            }
            else {
                updateStatus(result.status);
            }
    });
}
```

本機函式 `getSlice` 會呼叫 **File.getSliceAsync** 方法，從 **File** 物件中擷取配量。**getSliceAsync** 方法從配量集合中傳回 **Slice** 物件。它有兩個必要參數，_sliceIndex_ 和 _callback_。_sliceIndex_ 參數將整數作為配量集合中的索引子。如同 JavaScript API for Office 中的其他函式，**getSliceAsync** 方法也會將回呼函式作為參數來處理方法呼叫的結果。

**Slice** 物件可讓您存取檔案中包含的資料。除非 _getFileAsync_ 方法的 **options** 參數另有指定，否則 **Slice** 物件的大小是 4 MB。**Slice** 物件公開三個屬性︰[size](../../reference/shared/slice.size.md)、[data](../../reference/shared/slice.data.md) 和 [index](../../reference/shared/slice.index.md)。**size** 屬性會取得配量的大小，以位元組為單位。**index** 屬性會取得整數，表示配量集合的配量位置。




```js

// Get a slice from the file and then call sendSlice.
function getSlice(state) {

    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {

            updateStatus("Sending piece " + (state.counter + 1) +
                " of " + state.sliceCount);

            sendSlice(result.value, state);
        }
        else {
            updateStatus(result.status);
        }
    });
}
```

**Slice.data** 屬性會傳回檔案的原生資料以作為位元組陣列。如果資料是文字格式 (也就是 XML 或純文字)，配量即包含原生文字。如果您傳入 **Document.getFileAsync** 的 _fileType_ 參數的 **Office.FileType.Compressed**，配量會包含檔案的二進位資料來作為位元組陣列。在 PowerPoint 或 Word 檔案案例中，配量包含位元組陣列。

您必須實作自己的函式 (或使用可用的程式庫)，將位元組陣列資料轉換為 Base64 編碼的字串。有關如何使用 JavaScript 進行 Base64 編碼的詳細資訊，請參閱 [Base64 編碼和解碼](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding)。

一旦將資料轉換成 Base64，接著可以多種方法將它傳輸至網頁伺服器，包括作為 HTTP POST 要求的本文。

加入下列程式碼將配量傳送到網頁服務。


 >**附註：**此程式碼會將 PowerPoint 或 Word 檔案在多個快訊中傳送到 Web 伺服器。網頁伺服器或服務必須先將每個個別的配量編譯為單一的 .pptx 檔案，才能夠在其上執行任何操作。




```js

function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't 
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/en-US/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request 
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status 
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                }
                else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST 
        // request to the web server.
        request.send(fileData);
    }
}
```



正如其名，**File.closeAsync** 方法會關閉文件的連線並釋放資源。雖然 Office 增益集沙箱會收集檔案的參考範圍外的垃圾，但最佳作法是在程式碼完成工作時明確關閉檔案。**closeAsync** 方法具有單一參數 _callback_，可指定在完成呼叫時要呼叫的函式。




```js

function closeFile(state) {

    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            updateStatus("File closed.");
        }
        else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```
