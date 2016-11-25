
# <a name="context.document-property"></a>Context.document 屬性
取得代表與增益集互動之文件的物件。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```js
var _document = Office.context.document;
```


## <a name="return-value"></a>傳回值

[Document](../../reference/shared/document.md) 物件。


## <a name="remarks"></a>備註

增益集可以使用  **document** 屬性存取 API，和文件、活頁簿、簡報、專案和資料庫 (在 Access web 應用程式中) 的內容互動。


## <a name="example"></a>範例




```js
// Extension initialization code.
var _document;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Initialize instance variables to access API objects.
    _document = Office.context.document;
    });
}

```


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此屬性。空白儲存格表示 Office 主應用程式不支援此屬性。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增支援讓 **Office.context.document** 在 Access 的內容增益集中存取資料庫。|
|1.0|已導入|