# <a name="office-common-api-requirement-sets"></a>Office 通用 API 需求集合

需求集合是 API 成員的具名群組。Office 增益集使用資訊清單中所指定的需求集合，或使用執行階段檢查，以判定 Office 主應用程式是否支援增益集所需的的 API。如需詳細資訊，請參閱[指定 Office 主應用程式及 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

如需 Office 主應用程式在何處支援增益集的詳細資訊，請參閱 [Office 增益集主應用程式和平台可用性](https://dev.office.com/add-in-availability)。

## <a name="host-specific-api-requirement-sets"></a>主應用程式特定 API 需求集合

如需有關 Excel、Word、OneNote、Outlook 及對話方塊 API 需求集合的詳細資訊，請參閱︰

- [Excel JavaScript API 需求集合](excel-api-requirement-sets.md) (ExcelApi)
- [Word JavaScript API 需求集](word-api-requirement-sets.md) (WordApi)
- [OneNote JavaScript API 需求集合](onenote-api-requirement-sets.md) (OneNoteApi)
- [了解 Outlook API 需求集合](../outlook/tutorial-api-requirement-sets.md) (MailBox)
- [對話方塊 API 需求集合](dialog-api-requirement-sets.md) (DialogApi)

## <a name="common-api-requirement-sets"></a>通用 API 需求集合

下表列出通用 API 需求集合、每個集合中的方法以及支援該需求集合的 Office 主應用程式。所有這些 API 需求集合都是 1.1 版。


|  需求集合  |  Office 主應用程式  |  集合中的方法  |
|:-----|-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint&nbsp;Online|Document.getActiveViewAsync|
| BindingEvents  | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | PowerPoint<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>Excel Online<br/>PowerPoint Online|支援以 Office Open XML (OOXML) 格式輸出成為位元組陣列<br>(Office.FileType.Compressed) 使用 Document.getFileAsync 方法時。|
| CustomXmlParts    | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DocumentEvents    | Excel<br>Excel Online<br>PowerPoint Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| File  | PowerPoint<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、<br>Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為 HTML (Office.CoercionType.Html)。|
| ImageCoercion | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.setSelectedDataAsync 方法寫入資料時，支援轉換為影像 (Office.CoercionType.Image)。|
| 信箱   |Windows 版 Outlook<br>Web 版 Outlook<br>Mac 版 Outlook<br>Outlook Web App |請參閱[了解 Outlook API 需求集合](./outlook/tutorial-api-requirement-sets.md)|
| MatrixBindings    | Excel<br>Excel Online<br>Word<br>Word Online|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為 ”matrix” (陣列的陣列) 資料結構 (Office.CoercionType.Matrix)。|
| OoxmlCoercion | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為 Open Office XML (OOXML) 格式 (Office.CoercionType.Ooxml)。|
| PartialTableBindings  | Access Web Apps||
| PdfFile   | PowerPoint<br/>PowerPoint Online<br/>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getFileAsync 方法時，支援<br>輸出為 PDF 格式 (Office.FileType.Pdf)。|
| Selection | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Access Web Apps<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為 ”table” 資料結構 (Office.CoercionType.able)。|
| TextBindings  | Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為文字格式 (Office.CoercionType.Text)。|
| TextFile  | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>|使用 Document.getFileAsync 方法時，支援輸出為文字格式 (Office.FileType.Text)。|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>不屬於需求集合一部分的方法

下列 Office 版 JavaScript API 中的方法不屬於需求集合的一部分。如果增益集需要其中任何一種方法，請使用增益集資訊清單中的 **方法** 和 **方法** 元素，以宣告這些元素是必要的，或使用 if 陳述式執行階段檢查。如需詳細資訊，請參閱[指定 Office 主應用程式和 API 需求](../docs/overview/specify-office-hosts-and-api-requirements.md)。

|**方法名稱**|**Office 主應用程式支援**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access Web App、 Excel 及 Excel Online|
|Document.getFilePropertiesAsync|Excel、Excel Online、Word 及 PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013 與 Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 與 Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 與 Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 與 Project Professional 2013|
|Document.getSelectedViewAsync|PowerPoint 與 PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 與 Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 與 Project Professional 2013|
|Document.goToByIdAsync|Excel、Excel Online、Word 及 PowerPoint|
|Settings.addHandlerAsync|Access Web App、Excel、Excel Online、Word 及 PowerPoint|
|Settings.refreshAsync|Access Web App、Excel、Excel Online、Word、PowerPoint 及 PowerPoint Online|
|Settings.removeHandlerAsync|Access Web App、Excel、Excel Online、Word 及 PowerPoint|
|TableBinding.clearFormatsAsync|Excel、Excel Online|
|TableBinding.setFormatsAsync|Excel、Excel Online|
|TableBinding.setTableOptionsAsync|Excel、Excel Online|

## <a name="additional-resources"></a>其他資源

- [指定 Office 主應用程式和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office 增益集的 XML 資訊清單](../../docs/overview/add-in-manifests.md)




