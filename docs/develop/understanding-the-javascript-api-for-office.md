
# <a name="understanding-the-javascript-api-for-office"></a>了解適用於 Office 的 JavaScript API



本文提供適用於 Office 的 JavaScript API 和如何使用它的相關資訊。如需參考資訊，請參閱[適用於 Office 的 JavaScript API](../../reference/javascript-api-for-office.md)。如需將 Visual Studio 專案檔更新至最新版的適用於 Office 的 JavaScript API 的相關資訊，請參閱[更新您的適用於 Office 的 JavaScript API 和資訊清單的結構描述檔案的版本](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。

>**附註：**建立增益集時，如果您打算[發佈](../publish/publish.md)增益集至 Office 市集中，請確定您符合 [Office 市集驗證原則](https://msdn.microsoft.com/en-us/library/jj220035.aspx)。例如，若要通過驗證，增益集必須可以在所有支援您定義的方法的平台上運作 (如需詳細資料，請參閱 [4.12 節](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3)與 [Office 增益集主應用程式與可用性頁面](https://dev.office.com/add-in-availability))。

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>在增益集中參考適用於 Office 程式庫的 JavaScript API

[JavaScript API for Office](../../reference/javascript-api-for-office.md) 程式庫包含 Office.js 檔案和關聯的主應用程式特定的 .js 檔案，例如 Excel-15.js 和 Outlook-15.js。參考 API 的最簡單方法是藉由將下列 `<script>` 新增至您的頁面的 `<head>` 標記，使用我們的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

這將在增益集第一次載入時下載並快取適用於 Office 的 JavaScript API 檔案，以確定它使用 Office.js 的最新實作和指定的版本其相關聯的檔案。

如需 Office.js CDN 的詳細資訊，包括如何處理版本設定和回溯相容性，請參閱[從其內容傳遞網路 (CDN) 參考適用於 Office 的 JavaScript API](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

## <a name="initializing-your-add-in"></a>初始化增益集


 **適用於：**所有增益集類型


Office.js 能提供初始化事件，該事件會在 API 完全載入並準備好要開始與使用者互動時引發。您可以使用 **initialize** 事件處理常式來實作常見的增益集初始化案例，例如在 Excel 中提示使用者選取部份儲存格，然後插入以這些選取值初始化的圖表。您也可以使用初始化事件處理常式，為增益集初始化其他的自訂邏輯，例如建立繫結、提示輸入預設的增益集設定值等等。

 初始化事件看起來至少如下列範例所示︰     

```js
Office.initialize = function () { };
```
如果您使用包含自己的初始化處理常式或測試的其他 JavaScript 架構，則這些項目應該放在 Office.initialize 事件中。例如，[JQuery](https://jquery.com) 之 `$(document).ready()` 函式的參考方式如下：

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```
將事件處理常式指派給初始化事件 **Office.initialize** 需要 Office 增益集內的所有網頁。如果您未指派事件處理常式，增益集啟動時可能會引發錯誤。此外，如果使用者嘗試搭配 Office Online 網頁用戶端 (例如 Excel Online、PowerPoint Online 或 Outlook Web App) 使用您的增益集，它將無法執行。如果您不需要任何初始化程式碼，指派給 **Office.initialize** 的函式主體可以是空的，如前文第一個範例所示。

如需初始化增益集時事件順序的詳細資訊，請參閱[載入 DOM 和執行階段環境](../../docs/develop/loading-the-dom-and-runtime-environment.md)。

#### <a name="initialization-reason"></a>初始化原因
對於執行窗格和內容增益集，Office.initialize 提供了額外 _reason_ 參數。這個參數可以用來判斷如何將增益集新增至目前的文件中。您可以使用這個參數，對於增益集第一次插入與增益集已存在於文件中，提供不同的邏輯。 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
      switch (reason) {
        case 'inserted': console.log('The add-in was just inserted.');
        case 'documentOpened': console.log('The add-in is already part of the document.');
    }
}
```
如需詳細資訊，請參閱 [Office.initialize 事件](../../reference/shared/office.initialize.md)和 [InitializationReason 列舉](../../reference/shared/initializationreason-enumeration.md) 

## <a name="context-object"></a>內容物件

 **適用於：**所有增益集類型

初始化增益集時，它會有許多可以在執行階段環境中與其互動的不同物件。增益集的執行階段內容會透過 API 的[Context](../../reference/shared/office.context.md) 物件反映。**Context** 是主要的物件，提供 API 最重要物件的存取，例如 [Document](../../reference/shared/document.md) 和 [Mailbox](../../reference/outlook/Office.context.mailbox.md) 物件，其分別提供文件及信箱內容的存取。

例如，在工作窗格或內容增益集中，您可以使用 [Context](../../reference/shared/office.context.document.md) 物件的 **document** 屬性來存取 **Document** 物件的屬性和方法，以與 Word 文件、Excel 工作表或 Project 的排程內容互動。同樣地，在 Outlook 增益集中，您可以使用 [Context](../../reference/outlook/Office.context.mailbox.md) 物件的 **mailbox** 屬性來存取 **Mailbox** 物件的屬性和方法，以與郵件、會議邀請或約會內容互動。

**Context** 物件也會提供 [contentLanguage](../../reference/shared/office.context.contentlanguage.md) 和 [displayLanguage](../../reference/shared/office.context.displaylanguage.md) 屬性的存取，可讓您決定文件或項目中或主應用程式使用的地區設定 (語言)。而且，[roamingSettings](../../reference/outlook/Office.context.md) 屬性可讓您存取 [RoamingSettings](../../reference/outlook/RoamingSettings.md) 物件的成員。最後，**Context** 物件提供可讓增益集啟動快顯對話方塊的 [ui](../../reference/shared/officeui.md) 屬性。


## <a name="document-object"></a>Document 物件


 **適用於：**內容和工作窗格增益集類型

為了與 Excel、PowerPoint 和 Word 中的文件資料互動，API 提供了 [Document](../../reference/shared/document.md) 物件。您可以使用 **Document** 物件成員，透過下列方式存取資料︰


- 以文字、連續的儲存格 (矩陣) 或表格形式讀取和寫入作用中選取範圍。
    
- 表格式資料 (矩陣或表格)。
    
- 繫結 (使用 **Bindings** 物件的 "add" 方法建立)。
    
- 自訂 XML 組件 (僅適用於 Word)。
    
- 在文件上每個增益集保存的設定或增益集狀態。
    
您也可以使用 **Document** 物件來與 Project 文件中的資料互動。API 的 Project 特定功能會在成員 [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) 抽象類別中詳加說明。如需建立 Project 的工作窗格增益集的詳細資訊，請參閱 [Project 的工作窗格的增益集](../project/project-add-ins.md)。

這所有形式的資料存取都是從抽象的 **Document** 物件執行個體開始。

您可以藉由使用 **Context** 物件的 [document](../../reference/shared/office.context.document.md) 屬性，在初始化工作窗格或內容增益集時存取 **Document** 物件的執行個體。**Document** 物件會定義在 Word 和 Excel 文件間共用的通用資料存取函式，並也提供 Word 文件的 **CustomXmlParts** 物件的存取。

**Document** 物件支援讓開發人員可以存取文件內容的四種方法︰


- 以選項為基礎的存取
    
- 以繫結為基礎的存取
    
- 自訂 XML 組件式存取 (僅限 Word)
    
- 整個文件式的存取 (PowerPoint 和 Word)
    
為了幫助您了解以選取範圍和繫結為基礎的資料存取方法如何運作，我們將先說明資料存取 API 如何跨不同的 Office 應用程式提供一致的資料存取。


### <a name="consistent-data-access-across-office-applications"></a>所有 Office 應用程式一致的資料存取

 **適用於：**內容和工作窗格增益集類型

若要建立順暢地跨不同的 Office 文件運作的擴充功能，適用於 Office 的 JavaScript API 會透過常見的資料類型取出每個 Office 應用程式的特殊性，並且能夠將不同的文件內容強制轉型為三個常見的資料類型。


#### <a name="common-data-types"></a>常見的資料類型

在這兩個以選取範圍為基礎和以繫結為基礎的資料存取中，文件內容是透過所有支援的 Office 應用程式之間的通用資料類型公開。Office 2013 中支援三種主要的資料類型︰



|**資料類型**|**描述**|**主機應用程式支援**|
|:-----|:-----|:-----|
|文字|在選取範圍或繫結中提供資料的字串表示。|在 Excel 2013、Project 2013和 PowerPoint 2013 中，僅支援使用純文字。在 Word 2013 中，支援三種文字格式︰純文字、HTML 和 Office Open XML (OOXML)。在 Excel 中選取儲存格中的文字時，選取項目為主的方法會讀取和寫入儲存格的整個內容，即使只有在儲存格中選取一部分的文字。在 Word 和 PowerPoint 中選取文字時，以選項為基礎的方法僅會讀取和寫入至已選取的字元執行。Project 2013 和 PowerPoint 2013 僅支援以選項為基礎的資料存取。|
|矩陣|以兩個維度**陣列**提供選取範圍或繫結中的資料，在 JavaScript 會實作為陣列的陣列。例如，兩欄中兩列的 **string** 值會是 ` [['a', 'b'], ['c', 'd']]`，而三列的單一欄會是 `[['a'], ['b'], ['c']]`。|只有在 Excel 2013 和 Word 2013 中才支援矩陣資料存取。|
|表格|在選取範圍或繫結中提供資料做為 [TableData](../../reference/shared/tabledata.md) 物件。**TableData** 物件會透過**標頭**和**列**屬性公開資料。|只有在 Excel 2013 和 Word 2013 中支援表格資料存取。|

#### <a name="data-type-coercion"></a>資料類型強制型轉

**Document** 和 [Binding](../../reference/shared/binding.md) 物件上的資料存取方法，支援使用這些方法的 _coercionType_ 參數及對應的 [CoercionType](../../reference/shared/coerciontype-enumeration.md) 列舉值來指定所需的資料類型。無論繫結的實際的形狀為何，不同的 Office 應用程式會透過嘗試將資料強制轉型為要求的資料類型來支援常見的資料類型。例如，如果選取了 Word 表格或段落時，開發人員可以指定以純文字、HTML、Office Open XML 或表格讀取它，而 API 實作會處理所需的轉換和資料轉換。


 >**提示：**   **何時應使用矩陣與資料表 coercionType 進行資料存取？**如果您需要表格式資料在加入列及欄時動態成長，而您必須處理表格標題，應該使用表格式資料類型 (藉由將 **Document** 或 **Binding** 物件資料存取方法的 _coercionType_ 參數指定為 `"table"` 或 **Office.CoercionType.Table**)。於資料結構內加入列及欄在資料表和矩陣資料中支援，但附加列及欄僅對表格資料支援。如果您不打算加入列及欄，並且您的資料並不需要標頭功能，那麼您應該使用矩陣資料類型 (藉由將資料存取方法的 _coercionType_ 參數指定為 `"matrix"` 或 **Office.CoercionType.Matrix**)，提供與資料互動的更簡單模型。

如果無法將資料強制轉型為指定的類型，回撥中的 [AsyncResult.status](../../reference/shared/asyncresult.error.md) 屬性會傳回 `"failed"`，而您可以使用 [AsyncResult.error](../../reference/shared/asyncresult.context.md) 屬性來存取含有方法呼叫失敗原因相關資訊的 [Error](../../reference/shared/error.md) 物件。


## <a name="working-with-selections-using-the-document-object"></a>使用 Document 物件處理選取範圍


**Document** 物件公開的方法可讓您以「取得並忘記」方式來讀取和寫入至使用者目前的選取範圍。若要這樣做，**Document** 物件可提供 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法。

如需示範如何利用選取項目來執行工作的程式碼範例，請參閱[在文件或試算表中的作用選取範圍內讀取和寫入資料](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。


## <a name="working-with-bindings-using-the-bindings-and-binding-objects"></a>處理使用 Bindings 和 Binding 物件的繫結


以繫結為基礎的資料存取，可讓內容和工作窗格增益集一致地透過與繫結相關聯的識別碼來存取文件或試算表的特定區域。增益集必須先藉由呼叫將文件的一部分與唯一識別碼產生關聯的其中一個方法來建立繫結︰[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)、[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) 或 [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)。建立繫結之後，增益集可以使用提供的識別碼來存取文件或試算表的關聯的區域中所包含的資料。建立繫結可提供增益集下列值︰


- 允許存取所有支援的 Office 應用程式的常見資料結構，例如︰表格、範圍或文字 (連續執行的字元)。
    
- 啟用讀寫作業而不需要使用者進行選擇。
    
- 在文件中的增益集和資料之間建立關聯。繫結會保存在文件中，而且可以稍後加以存取。
    
建立繫結也可讓您訂閱資料及選取範圍變更事件，其範圍限制在文件或試算表的特定區域。這表示增益集只會收到繫結區域內發生的變更的通知，而不是整份文件或試算表一般變更的通知。

[Bindings](../../reference/shared/bindings.bindings.md) 物件會公開 [getAllAsync](../../reference/shared/bindings.getallasync.md) 方法，其可為文件或試算表上建立的所有繫結的集合提供存取。可以使用 [Bindings.getBindingByIdAsync](../../reference/shared/bindings.getbyidasync.md) 或 [Office.select](../../reference/shared/office.select.md) 方法利用 ID 存取個別的繫結。您可以使用 **Bindings** 物件的下列其中一個方法來建立新繫結，以及移除現有的繫結︰[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)、[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)、[addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md) 或 [releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)。

當您使用 _addFromSelectionAsync_、**addFromPromptAsync** 或 **addFromNamedItemAsync** 方法建立繫結時，對 **bindingType** 參數可以指定三種不同類型的繫結︰



|**繫結類型**|**描述**|**主機應用程式支援**|
|:-----|:-----|:-----|
|文字繫結|繫結至可以以文字表示之文件的區域。|在 Word 中，大部分的連續選取項目為有效，而在 Excel 中，只有單一儲存格選取項目可以是文字繫結的目標。在 Excel 中僅支援純文字。Word 中支援三種格式︰純文字、HTML 和 Open XML for Office。|
|矩陣繫結|繫結至包含表格式資料而沒有標題之文件的固定區域。矩陣繫結中的資料會以兩個維度**陣列**的形式寫入或讀取，在 JavaScript 會實作為陣列的陣列。例如，兩欄中兩列的 **string** 值可以寫入或讀取為 ` [['a', 'b'], ['c', 'd']]`，而三列的單一欄可以寫入或讀取為 `[['a'], ['b'], ['c']]`。|在 Excel 中，任何連續選取的儲存格可以用於建立矩陣繫結。在 Word 中，只有表格支援矩陣繫結。|
|表格繫結|繫結至包含標題之表格文件的某個區域。在表格繫結中的資料會以 [TableData](../../reference/shared/tabledata.md) 物件的形式寫入或讀取。**TableData** 物件會透過**標頭**和**列**屬性公開資料。|任何 Excel 或 Word 表格可以是表格繫結的基礎。建立表格繫結之後，使用者加入至表格的每個新列或欄會自動包含在繫結中。 |
使用 **Bindings** 物件的三個 "add" 方法其中一個建立繫結之後，您可以使用對應的物件的方法來處理繫結的資料和屬性︰[MatrixBinding](../../reference/shared/binding.matrixbinding.md)、[TableBinding](../../reference/shared/binding.tablebinding.md) 或 [TextBinding](../../reference/shared/binding.textbinding.md)。這三個物件皆繼承可讓您與繫結資料互動的 [Binding](../../reference/shared/binding.getdataasync.md) 物件的 [getDataAsync](../../reference/shared/binding.setdataasync.md) 和 **setDataAsync** 方法。

如需示範如何使用繫結執行工作的程式碼範例，請參閱[繫結至文件或試算表中的區域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="working-with-custom-xml-parts-using-the-customxmlparts-and-customxmlpart-objects"></a>處理使用 CustomXmlParts 和 CustomXmlPart 物件的自訂 XML 組件


 **適用於：**Word 的工作窗格增益集

API 的 [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md) 和 [CustomXmlPart](../../reference/shared/customxmlpart.customxmlpart.md) 物件可提供 Word 文件中的自訂 XML 組件的存取，其可啟用文件內容 XML 導向的操作。如需使用 **CustomXmlParts** 和 **CustomXmlPart** 物件的示範，請參閱 [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) 程式碼範例。


## <a name="working-with-the-entire-document-using-the-getfileasync-method"></a>使用 getFileAsync 方法處理整份文件


 **適用於：**Word 和 PowerPoint 的工作窗格增益集

[Document.getFileAsync](../../reference/shared/document.getfileasync.md) 方法和 [File](../../reference/shared/file.md) 與 [Slice](../../reference/shared/slice.md) 物件的成員，可提供以一次最多 4 MB 切片 (區塊) 的形式，取得整個 Word 和 PowerPoint 文件檔案的功能。如需相關資訊，請參閱 [How To：從增益集中的文件取得所有檔案內容](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)。


## <a name="mailbox-object"></a>信箱物件


 **適用於：**Outlook 增益集

Outlook 增益集主要是使用透過 [Mailbox](../../reference/outlook/Office.context.mailbox.md) 物件公開的 API 子集。若要存取物件和成員，特別是針對用於 Outlook 增益集，例如 [Item](../../reference/outlook/Office.context.mailbox.item.md) 物件，您會使用 [Context](../../reference/outlook/Office.context.mailbox.md) 物件的 **mailbox** 屬性來存取 **Mailbox** 物件，如下列程式碼所示。




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

此外，Outlook 增益集可以使用下列物件︰


-  **Office** 物件︰用於初始化。
    
-  **Context** 物件︰針對內容和顯示語言屬性的存取。
    
-  **RoamingSettings** 物件︰用於將 Outlook 增益集特定的自訂設定儲存至安裝增益集所在使用者的信箱。
    
如需在 Outlook 增益集中使用 JavaScript 的詳細資訊，請參閱 [Outlook 增益集](../outlook/outlook-add-ins.md)和 [Outlook 增益集的架構和功能概觀](../outlook/overview.md)。


## <a name="api-support-matrix"></a>API 支援矩陣


此表格摘要說明 API 和跨增益集類型 (內容、工作窗格和 Outlook) 所支援的功能，以及可主控它們當您使用 [1.1 增益集資訊清單結構描述和 1.1 版適用於 Office 的 JavaScript API 支援的功能](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx)指定[您的增益集支援的 Office 應用程式](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)時，可以主控它們的 Office 應用程式。


|||||||||
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
||**主應用程式名稱**|資料庫|活頁簿|信箱|Presentation|Document|Project|
||**支援的****主應用程式**|Access Web App|Excel Online|Outlook Web App OWA for Devices|PowerPoint Online|Word|Project|
|**支援的增益集類型**|內容|是|Y||是|||
||工作窗格||是||Y|Y|是|
||Outlook|||是||||
|**支援的 API 功能**|讀寫文字||是||Y|是|是 (唯讀)|
||讀寫矩陣||是|||是||
||讀寫資料表||是|||是||
||讀寫 HTML|||||是||
||讀寫 Office Open XML|||||是||
||讀取工作、資源、檢視和欄位屬性||||||是|
||選取範圍變更事件||是|||是||
||取得整份文件||||是|是||
||繫結和繫結事件|是 (僅完整和部分表格繫結)|是|||是||
||讀寫自訂 XML 組件|||||是||
||保存增益集狀態資料 (設定)|是 (每一主應用程式增益集)|是 (每份文件)|是 (每個信箱)|是 (每份文件)|是 (每份文件)||
||設定變更的事件|是|Y||Y|是||
||取得使用中檢視模式和檢視變更的事件||||是|||
||導覽至文件中的位置||是||Y|是||
||使用規則和 RegEx 啟動內容|||是||||
||讀取項目內容|||是||||
||讀取使用者設定檔|||是||||
||取得附件|||是||||
||取得使用者識別權杖|||是||||
||呼叫 Exchange Web 服務|||是||||
