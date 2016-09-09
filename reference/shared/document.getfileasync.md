
# Document.getFileAsync 方法
傳回整份文件檔案，配量最多為 4194304 位元組 (4MB)。對於增益集的 iOS 檔案配量最多支援 65536 (64KB)。請注意，指定超過允許限制的檔案配量大小將造成「內部錯誤」失敗。 

|||
|:-----|:-----|
|**主機︰**|Excel、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|File|
|**上次變更於檔案**|1.1|

```js
Office.context.document.getFileAsync(fileType [, options], callback);
```


## 參數



|**名稱**|**類型	**|**說明**|**支援附註**|
|:-----|:-----|:-----|:-----|
| _fileType_|[FileType](../../reference/shared/filetype-enumeration.md)|指定將傳回檔案的格式。必要。<br/><table><tr><th>主應用程式</th><th>支援的 fileType</th></tr><tr><td>Excel Online</td><td>Office.FileType.Compressed</td></tr><tr><td>PowerPoint on Windows desktop</td><td>Office.FileType.Compressed、Office.FileType.Pdf</td></tr><tr><td>Word on Windows desktop、MAC 和 iPad</td><td>Office.FileType.Compressed、Office.FileType.Pdf、 Office.FileType.Text</td></tr><tr><td>Word Online</td><td>Office.FileType.Compressed、Office.FileType.Pdf、Office.FileType.Text</td></tr><tr><td>PowerPoint Online</td><td>Office.FileType.Compressed、Office.FileType.Pdf</td></tr></table>|**已變更於** 1.1，請參閱[支援歷程記錄](#支援歷程記錄)|
| _options_|**物件**|指定下列任何一項[選擇性參數](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _sliceSize_|**number**|指定您想要的配量大小 (以位元組計)，最多為 4194304 位元組 (4MB)。如果未指定，將使用 4194304 位元組 (4MB) 的預設配量大小。 ||
| _asyncContext_|**陣列**、**布林值**、**null**、**數字**、**物件**、**字串**或**未定義**|無變更的情況下，於 **AsyncResult** 物件中傳回的任一類型使用者定義項目。||
| _callback_|**物件**|回呼傳回時所叫用的函數，其唯一的參數為 **AsyncResult** 類型。||

## 回呼值

傳遞至 _callback_ 參數的函數執行時，該函數會收到 [AsyncResult](../../reference/shared/asyncresult.md) 物件，您可以從回呼函數的唯一參數存取該物件。

在傳遞至 **getFileAsync** 方法的回呼函數中，您可以使用 **AsyncResult** 物件的屬性以傳回下列資訊。



|**屬性**|**用途**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|存取 [File](../../reference/shared/file.md) 物件。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|判定作業成功或失敗。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|作業失敗時，存取提供錯誤資訊的 [Error](../../reference/shared/error.md) 物件。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|存取您的使用者定義**物件**或值 (如果您傳遞了其中一項做為 _asyncContext_ 參數)。|

## 備註

對於 Office for iOS 以外的 Office 主應用程式中執行的增益集，**getFileAsync** 方法支援取得最高 4194304 位元組 (4MB) 的檔案配量。對於 Office for iOS 應用程式中執行的增益集，**getFileAsync** 方法支援取得最高 65536 位元組 (64KB) 的檔案配量。

可使用以下列舉或文字值指定 _fileType_ 參數。


**FileType 列舉**


|**列舉**|**值**|**說明**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|以 Office Open XML (OOXML) 格式，傳回整個文件，成為位元組陣列。|
|Office.FileType.Pdf|"pdf"|以 PDF 格式傳回整個文件，成為位元組陣列。|
|Office.FileType.Text|"text"|只傳回文件的文字，成為**字串**。 |
不允許記憶體中有超過兩個文件；否則 **getFileAsync** 作業會失敗。當您完成使用檔案時，請使用 [File.closeAsync](../../reference/shared/file.closeasync.md)方法關閉檔案。


## 範例 - 以 Office Open XML (“compressed”) 格式取得文件

下列範例會以 Office Open XML (“compressed”) 格式，65536 位元組 (64KB) 配量，取得文件。附註：此範例中的 `app.showNotification` 實作是來自 Office 增益集的 Visual Studio 範本。


```js
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ }, 
        function (result) {
            if (result.status == "succeeded") {
            // If the getFileAsync call succeeded, then
            // result.value will return a valid File Object.
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

            // Get the file slices.
            getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
            else {
            app.showNotification("Error:", result.error.message);
            }
    });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
    file.getSliceAsync(nextSlice, function (sliceResult) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
               // All slices have been received.
               file.closeAsync();
               onGotAllSlices(docdataSlices);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
        }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
    });
}

function onGotAllSlices(docdataSlices) {
    var docdata = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    var fileContent = new String();
    for (var j = 0; j < docdata.length; j++) {
        fileContent += String.fromCharCode(docdata[j]);
    }

    // Now all the file content is stored in 'fileContent' variable,
    // you can do something with it, such as print, fax...
}

```


## 範例 - 以 PDF 格式取得文件

下列範例會以 PDF 格式取得文件。


```js
Office.context.document.getFileAsync(Office.FileType.Pdf,
    function(result) {
        if (result.status == "succeeded") {
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
            // Now, you can call getSliceAsync to download the files, as described in the previous code segment (compressed format).
            
            myFile.closeAsync();
        }
        else {
            app.showNotification("Error:", result.error.message);
        }
}
);


```


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|File|
|**最低權限等級**|[ReadAllDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.1| 在 PowerPoint Online 中，新增支援 **Office.FileType.Pdf** 做為 _fileType_ 參數。|
|1.1| 在 PowerPoint Online 中，新增支援 **Office.FileType.Compressed** 做為 _fileType_ 參數。|
|1.1| 在 Word Online 中，新增支援 **Office.FileType.Text** 做為 _fileType_ 參數。|
|1.1| 在 Excel Online 中，新增支援 **Office.FileType.Compressed** 做為 _fileType_ 參數。|
|1.1| 在 Word Online 中，新增支援 **Office.FileType.Compressed** 和 **Office.FileType.Pdf** 做為 _fileType_ 參數。|
|1.1|在 iPad 版 Office 的 PowerPointWord 中，新增支援所有 **FileType** 值做為 _fileType_ 參數。|
|1.1|在 Windows 桌面版 Word 和 PowerPoint 中，新增支援 **Office.FileType.Pdf** 做為 _fileType_ 參數。|
|1.0|已導入|
