
# Office.initialize 事件
已載入執行階段環境，且增益集準備好開始與裝載它的文件互動時，就會發生。 

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```js
Office.initialize = function (reason) {/* initialization code */}
```


## 備註

_initialize_ 事件接聽程式函數的 **reason** 參數會傳回 [InitializationReason](../../reference/shared/initializationreason-enumeration.md) 列舉值，指定初始化如何發生。有兩種方式可以初始化工作窗格或內容增益集︰


- 使用者剛在 Office 主應用程式功能區的 **[插入]** 索引標籤上，從 **[增益集]** 下拉式清單的 **[最近使用的增益集]** 區段將其插入，或是從 **[插入增益集]** 對話方塊將其插入。
    
- 使用者開啟已經包含增益集的文件。
    

 >**附註**：針對工作窗格和內容增益集，**initialize** 事件接聽程式函數的 reason 參數只會傳回 **InitializationReason** 列舉值。它不會針對 Outlook 增益集傳回值。


## 範例

您可以使用 **InitializationEnumeration** 的值，針對增益集是第一次插入，以及它已經是文件一部分的狀況，實作不同的邏輯。下列範例會示範一些簡單的邏輯，使用 _reason_ 參數的值以顯示工作窗格或內容增益集初始化的方式。


```js
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Display initialization reason.
    if (reason == "inserted")
    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此事件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Mac 版 Outlook**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|是|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、Outlook、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增對於初始化 Access 內容增益集的支援。|
|1.0|已導入|
