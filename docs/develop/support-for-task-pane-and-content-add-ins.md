
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Office 2013 中內容和工作窗格增益集的 Office JavaScript API 支援


您可以使用 [Office JavaScript API](../../reference/javascript-api-for-office.md) 來建立 Office 2013 主應用程式的工作窗格或內容增益集。內容和工作窗格增益集所支援的物件與方法分類如下︰


1. **與其他 Office 增益集共用的通用物件。**這些物件包含 [Office](../../reference/shared/office.md)、[Context](../../reference/shared/office.context.md) 和 [AsyncResult](../../reference/shared/asyncresult.md)。**Office** 物件是 Office JavaScript API 的根物件。**Context** 物件代表增益集執行階段環境。**Office** 和 **Context** 兩者為所有 Office 增益集的基本物件。**AsyncResult** 物件代表非同步作業的結果，例如要傳回至 **getSelectedDataAsync** 方法的資料，該方法是讀取使用者在文件中所選取的內容。
    
2.  **Document 物件。**大部分內容和工作窗格增益集可用的 API 已透過 [Document](../../reference/shared/document.md) 物件的方法、屬性和事件公開。內容或工作窗格增益集可以使用 [Office.context.document](../../reference/shared/office.context.document.md) 屬性來存取 **Document** 物件，且透過它，便可存取 API 的關鍵成員，例如 [Bindings](../../reference/shared/bindings.bindings.md) 和 [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md) 物件，以及 [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)、[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)，和 [getFileAsync](../../reference/shared/document.getfileasync.md) 方法，來使用文件中的資料。**Document** 物件也提供 [mode](../../reference/shared/document.mode.md) 屬性來判斷文件是否唯讀或處於編輯模式，[url](../../reference/shared/document.url.md) 屬性可讓您取得目前文件的 URL，並存取 [Settings](../../reference/shared/settings.md) 物件。**Document** 物件也支援加入 [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) 事件的事件處理常式，您就可以偵測到使用者何時變更其在文件中的選取範圍。
    
   內容或工作窗格增益集只能在載入 DOM 和執行階段環境後，才能存取 **Document** 物件，通常位於 [Office.initialize](../../reference/shared/office.initialize.md) 事件的事件處理常式中。如需有關初始化增益集時事件流程的相關資訊，以及如何檢查 DOM 和執行階段環境是否成功載入，請參閱[載入 DOM 和執行階段環境](../../docs/develop/loading-the-dom-and-runtime-environment.md)。
    
3.  **使用特定功能的物件。**若要使用 API 的特定功能，請使用下列物件和方法︰
    
    - [Bindings](../../reference/shared/bindings.bindings.md) 物件的方法可建立或取得繫結；[Bindings](../../reference/shared/binding.md) 物件的方法和屬性可使用資料。
    
    - [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md)、[CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) 和關聯物件可建立及管理 Word 文件中的自訂 XML 組件。
    
    - [File](../../reference/shared/file.md) 和 [Slice](../../reference/shared/slice.md) 物件可建立整份文件，將它分割成區塊或「切片」，然後讀取或傳輸這些切片中的資料。
    
    - [Settings](../../reference/shared/settings.md) 物件可儲存自訂資料，例如使用者喜好設定以及增益集狀態。
    

 >**重要事項**  在主控內容和工作窗格增益集的所有 Office 應用程式之間不受支援的某些 API 成員。若要判斷哪些成員受到支援，請參閱下列其中一項︰

如需 Office 主應用程式之間 Office JavaScript API 支援的摘要，請參閱[瞭解 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)。


## <a name="reading-and-writing-to-an-active-selection"></a>讀取和寫入使用中的選取範圍

您可以讀取或寫入文件、試算表或簡報中使用者目前的選取範圍。根據增益集的主應用程式，您可以指定資料結構的類別，來讀取或寫入作為 [Document](../../reference/shared/document.getselecteddataasync.md) 物件的 [getSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) 和 [setSelectedDataAsync](../../reference/shared/document.md) 方法中的參數。例如，您可以在 Word 指定任何類型的資料 (文字、HTML、表格式資料或 Office Open XML)、在 Excel 指定文字和表格式資料，以及在 PowerPoint 和 Project 指定文字。您也可以建立事件處理常式，即可偵測對使用者選取範圍所做的變更。下列範例會使用 **getSelectedDataAsync** 方法從作為文字的選取範圍取得資料。


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

如需更多詳細資料和範例，請參閱[在文件或試算表中將資料讀取和寫入使用中的選取範圍](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>繫結至文件或試算表中的區域

您可以使用  **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法來讀取或寫入文件、試算表或簡報中使用者*目前*的選取範圍。不過，如果您想要存取跨區域的文件中，執行增益集的相同區域，而不需使用者先進行選取，您應該先繫結至該區域。您也可以訂閱該繫結區域的資料與選取範圍變更事件。

您可以使用 [Bindings](../../reference/shared/bindings.addfromnameditemasync.md) 物件的 [addFromNamedItemAsync](../../reference/shared/bindings.addfrompromptasync.md)、[addFromPromptAsync](../../reference/shared/bindings.addfromselectionasync.md)，或 [addFromSelectionAsync](../../reference/shared/bindings.bindings.md) 方法來新增繫結。這些方法會傳回識別碼，您可以使用該識別碼來存取繫結中的資料，或訂閱其資料變更或選取範圍變更事件。

下列範例是將繫結新增至文件中目前選取的文字，方法是使用 **Bindings.addFromSelectionAsync** 方法。



```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

如需詳細資訊和範例，請參閱[繫結至文件或試算表中的區域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="getting-entire-documents"></a>取得整個文件

如果工作窗格增益集在 PowerPoint 或 Word 中執行，您可以使用 [Document.getFileAsync](../../reference/shared/document.getfileasync.md)、[File.getSliceAsync](../../reference/shared/file.getsliceasync.md)，和 [File.closeAsync](../../reference/shared/file.closeasync.md) 方法以取得整個簡報或文件。

呼叫 **Document.getFileAsync** 時，您會得到一份 [File](../../reference/shared/file.md) 物件中的文件。**File** 物件讓您存取作為 [Slice](../../reference/shared/document.md) 物件表示之「區塊」中的文件。呼叫 **getFileAsync** 時，您可以指定檔案類型 (文字或壓縮的 Open Office XML 格式) 以及切片大小 (最多 4 MB)。若要存取 **File** 物件的內容，請呼叫 **File.getSliceAsync**，這會傳回 [Slice.data](../../reference/shared/slice.data.md) 屬性中的原始資料。如果您指定壓縮的格式，就會取得位元組陣列的檔案資料。如果您是將檔案傳輸至 web 服務，就可以在送出之前，將壓縮的原始資料轉換成 base64 編碼字串。最後，當您完成取得檔案的切片時，使用 **File.closeAsync** 方法來關閉文件。

如需詳細資訊，請參閱如何[從 PowerPoint 或 Word 增益集中，取得整份文件](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)。 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>讀取和寫入 Word 文件的自訂 XML 組件

使用 Open Office XML 檔案格式及內容控制項，可以將自訂 XML 組件加入至 Word 文件，並將 XML 組件中的元素繫結至該文件中的內容控制項。開啟文件時，Word 會讀取並自動將自訂 XML 組件的資料填入繫結內容控制項。使用者也可以將資料寫入至內容控制項，而當使用者儲存文件時，控制項中的資料將會儲存到繫結的 XML 組件。Word 的工作窗格增益集可以使用 [Document.customXmlParts](../../reference/shared/document.customxmlparts.md) 屬性、[CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md)、[CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md)，和 [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md) 物件，來動態讀取資料和將其寫入文件。

自訂 XML 組件可能與命名空間相關聯。若要從命名空間中的自訂 XML 組件中取得資料，請使用 [CustomXmlParts.getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md) 方法。

您也可以使用 [CustomXmlParts.getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md) 方法，透過其 Guid 來存取自訂 XML 組件。在取得自訂 XML 組件後，請使用 [CustomXmlPart.getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md) 方法來取得 XML 資料。

若要將新的自訂 XML 組件新增至文件，請使用  **Document.customXmlParts** 屬性來取得文件中的自訂 XML 組件，並呼叫 [CustomXmlParts.addAsync](../../reference/shared/customxmlparts.addasync.md) 方法。

如需如何使用工作窗格增益集的自訂 XML 組件的詳細資訊，請參閱[使用 Office Open XML 建立更出色的 Word 增益集](../../docs/word/create-better-add-ins-for-word-with-office-open-xml.md)。


## <a name="persisting-add-in-settings"></a>保存增益集設定


通常您需要儲存增益集的自訂資料，例如使用者的喜好設定或增益集的狀態，並存取下一次開啟增益集時的資料。您可以使用一般的 web 程式設計技術來儲存該資料，例如瀏覽器 cookie 或 HTML 5 web 存放區。或者，如果增益集在 Excel、PowerPoint 或 Word 中執行，您可以使用 [Settings](../../reference/shared/settings.md) 物件的方法。與 **Settings** 物件建立的資料會儲存在增益集已插入其中並一起儲存的試算表、簡報或文件中。此資料僅供建立其本身的增益集使用。

若要避免與文件儲存所在之伺服器的來回，在會執行階段的記憶體中管理使用 **Settings** 物件建立的資料。初始化增益集時會將先前儲存的設定資料載入到記憶體，以及當您呼叫 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法時，只會將對該資料所做的變更儲存回文件中。就內部而言，會將資料儲存在序列化 JSON 物件作為名稱/值組。您使用 [Settings](../../reference/shared/settings.get.md) 物件的 [get](../../reference/shared/settings.set.md)、[set](../../reference/shared/settings.removehandlerasync.md)，和 **remove** 方法，來讀取、寫入和刪除資料記憶體內複本的項目。下列程式碼示範如何建立名為 `themeColor` 的設定並將其值設定為「綠色」。




```js
Office.context.document.settings.set('themeColor', 'green');
```

因為使用 **set** 和 **remove** 方法建立或刪除的設定資料於資料記憶體內複本上作用，您必須呼叫 **saveAsync**，才能將對設定資料所做的變更保存到增益集正在使用的文件中。

如需有關透過 **Settings** 物件來使用自訂資料的詳細資訊，請參閱[保存增益集狀態與設定](../../docs/develop/persisting-add-in-state-and-settings.md)。


## <a name="reading-properties-of-a-project-document"></a>讀取專案文件的屬性

如果工作窗格在 Project 中執行增益集，增益集可以讀取使用中專案之部份專案欄位、資源及工作欄位的資料。若要這樣做，您使用 [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) 物件的方法和事件，其可擴充 **Document** 物件以提供其他 Project 特定的功能。

如需讀取 Project 資料的範例，請參閱[使用文字編輯器建立您第一個 Project 2013 的工作窗格增益集](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)。


## <a name="permissions-model-and-governance"></a>權限模型和控管

增益集使用其資訊清單中的  **Permissions** 元素來要求權限，便能從 Office JavaScript API 存取它所需要的功能等級。例如，如果增益集需要文件的讀取/寫入存取權，其資訊清單必須指定 `ReadWriteDocument` 作為其 **Permissions** 元素中的文字值。因為權限的存在是用來保護使用者的隱私權和安全性，最佳作法是您應該要求其功能所需的最低層級權限。下列範例示範如何要求工作窗格資訊清單中的 **ReadDocument** 權限。


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

如需詳細資訊，請參閱[要求用於內容和工作窗格增益集的 API 權限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)。


## <a name="additional-resources"></a>其他資源


- [Office JavaScript API](../../reference/javascript-api-for-office.md)
    
- 
  [Office 增益集資訊清單的結構描述參考](http://msdn.microsoft.com/en-us/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92.aspx)
    
- [疑難排解 Office 增益集的使用者錯誤](../../docs/testing/testing-and-troubleshooting.md)
    
