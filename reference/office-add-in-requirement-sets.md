
# Office 增益集需求集合

需求集合是 API 成員的具名群組。Office 增益集使用資訊清單中所指定的需求集合，或使用執行階段檢查，以判定 Office 主應用程式是否支援增益集所需的 API。如需詳細資訊，請參閱[指定 Office 主應用程式和 API 需求](../docs/overview/specify-office-hosts-and-api-requirements.md)。

若要廣泛了解 Office 主應用程式在何處支援增益集，請參閱 [Office 增益集主應用程式和平台可用性](https://dev.office.com/add-in-availability)頁面。

## 需求集合


下表列出需求集合的名稱、每個集合中的方法、支援該需求集合的 Office 主應用程式，以及 API 的版本號碼。

如需有關 Outlook 需求集合的資訊，請參閱[了解 Outlook API 需求集合](./outlook/tutorial-api-requirement-sets.md)。

|  集合名稱  |  Version  |  Office 主應用程式  |  集合中的方法  |
|:-----|-----|:-----|:-----|
| ExcelApi   | 1.2 | Excel 2016<br>Excel Online<br>iPad 版 Excel<br>|保護工作表<br>工作表函數<br>排序<br>篩選<br>R1C1 參照樣式<br>合併儲存格<br>調整列高及欄寬<br>Chart.getImage()<br>Range.getUsedRange(valuesOnly)|
| ExcelApi   | 1.1 | Excel 2016<br>Excel Online<br>iPad 版 Excel<br>|Excel 命名空間中的所有元素|
| WordApi    | 1.2 | Word 2016<br>Word 2016 for Mac<br>Word for iPad<br>Word Online (預覽) | Word 命名空間中的所有元素。 已在此 WordApi 版本中新增下列方法：<br>Body.select(selectionMode)<br>Body.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>contentControl.select(selectionMode)<br>contentControl.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>inlinePicture.paragraph<br>inlinePicture.delete<br>inlinePicture.insertBreak(breakType, insertLocation)<br>inlinePicture.insertFileFromBase64(base64file, insertLocation)<br>inlinePicture.insertHtml(html, insertLocation)<br>inlinePicture.insertInlinePictureFromBase64(base64file, insertLocation)<br>inlinePicture.insertOoxml(ooxml, insertLocation)<br>inlinePicture.insertParagraph(paragraphText, insertLocation)<br>inlinePicture.insertText(text, insertLocation)<br>inlinePicture.select(selectionMode)<br>paragraph.select(selectionMode)<br>range.inlinePictures<br>range.select(selectionMode)<br>range.insertInlinePictureFomBase64(base64EcodedImage, insertLocation)|
| WordApi    | 1.1 | Word 2016<br>Word 2016 for Mac<br>Word for iPad<br>|除了已新增至 WordApi 1.2 及更新版本的 API 成員以外，Word 命名空間的所有元素如上所列。|
| ActiveView | 1.1 | PowerPoint<br>PowerPoint Online|Document.getActiveViewAsync|
| BindingEvents  | 1.1 | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | 1.1 |PowerPoint<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>Excel Online<br/>PowerPoint Online|支援以 Office Open XML (OOXML) 格式輸出成為位元組陣列<br>(Office.FileType.Compressed) 使用 Document.getFileAsync 方法時。|
| CustomXmlParts    | 1.1 |Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogAPI | 1.1 | Excel<br>PowerPoint<br>Word 2016<br>Outlook|Office.context.ui.displayDialogAsync()<br>Office.context.ui.messageParent()<br>Office.context.ui.close()|
| DocumentEvents    | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| 檔案  | 1.1 | PowerPoint<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | 1.1 | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、<br>Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為 HTML (Office.CoercionType.Html)。|
| ImageCoercion | 1.1 | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.setSelectedDataAsync 方法寫入資料時，支援轉換為影像 (Office.CoercionType.Image)。|
| 信箱   |   | Windows 版 Outlook<br>Web 版 Outlook<br>Mac 版 Outlook<br>Outlook Web App |請參閱[了解 Outlook API 需求集合](./outlook/tutorial-api-requirement-sets.md)|
| MatrixBindings    | 1.1 | Excel<br>Excel Online<br>Word|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | 1.1 | Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為 ”matrix” (陣列的陣列) 資料結構 (Office.CoercionType.Matrix)。|
| OoxmlCoercion | 1.1 | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為 Open Office XML (OOXML) 格式 (Office.CoercionType.Ooxml)。|
| PartialTableBindings  | 1.1 | Access Web Apps||
| PdfFile   | 1.1 | PowerPoint<br/>PowerPoint Online<br/>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getFileAsync 方法時，支援<br>輸出為 PDF 格式 (Office.FileType.Pdf)。|
| Selection | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| 設定  | 1.1 | Access Web Apps<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | 1.1 | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | 1.1 | Access Web Apps<br>Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為 ”table” 資料結構 (Office.CoercionType.able)。|
| TextBindings  | 1.1 | Excel<br>Excel Online<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad|使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法讀取及寫入資料時，支援強制型轉為文字格式 (Office.CoercionType.Text)。|
| TextFile  | 1.1 | Word 2013 和更新版本<br>Word 2016 for Mac<br>Word Online<br>Word for iPad<br/>|使用 Document.getFileAsync 方法時，支援輸出為文字格式 (Office.FileType.Text)。|

## 不屬於需求集合一部分的方法


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

## 其他資源



- [指定 Office 主應用程式和 API 需求](../docs/overview/specify-office-hosts-and-api-requirements.md)

