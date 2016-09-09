
# Document.url 屬性
取得主應用程式目前已開啟的文件 URL。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Project、Word|
|**上次變更於**|1.1|

```
var docUrl = Office.context.document.url;
```


## 傳回值

文件 URL。如果 URL 無法使用，傳回 **null**。


## 備註

 **重要事項：****url** 屬性傳回資訊，可能在文件名稱中和儲存位置中包含個人識別資訊 (PII)。如果您必須儲存或傳輸此資訊，請務必以加密格式執行作業。


## 範例




```
function displayDocumentUrl() {
    write(Office.context.document.url);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄





****


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 Word Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增對 Access 內容增益集的支援。|
|1.0|已導入|
