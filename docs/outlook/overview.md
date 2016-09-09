
# Outlook 增益集的架構和功能概觀

由 XML 資訊清單和程式碼 (JavaScript 和 HTML) 所組成的 Outlook 增益集。資訊清單會指定增益集的名稱和描述，以及增益集整合到 Outlook 的方式。開發人員可以使用資訊清單，在命令介面上放置按鈕、連結關閉規則運算式符合項目等等。資訊清單也會定義用來裝載增益集的 JavaScript 和 HTML 程式碼的 URL。

當使用者或系統管理員取得增益集時，增益集的資訊清單會儲存到使用者的信箱或組織內。當 Outlook 啟動時，它會載入使用者已安裝的所有資訊清單、加以處理並設定增益集 (例如顯示命令介面中的按鈕，對目前所選的郵件執行規則運算式等等) 的所有擴充點。使用者現在可以使用增益集。

當使用者與增益集進行互動時，會從資訊清單中所指定的主機位置載入 JavaScript 和 HTML 檔案。

增益集會使用 Office.js API 來存取 Outlook 增益集 API 並與 Outlook 互動。


**當使用者啟動 Outlook 時，一般元件的互動**

![啟動 Outlook 郵件增益集的事件流程](../../images/olowawecon15_LoadingDOMAgaveRuntime.png)
### 版本設定

在我們發展 Outlook 用戶端和增益集平台並新增增益集的新方法來進行整合時，有時候我們無法同時跨所有的用戶端 (Mac、Windows、web、行動) 實作功能。為了處理這種情況，我們設定資訊清單和 API 的版本。使用這種方式，平台可在任何時間支援回溯相容性，這表示開發人員可以建置在較舊用戶端的下層方式中運作的增益集，並可讓您利用較新用戶端中的新功能。您可以在 [Outlook 增益集資訊清單](manifests/manifests.md)中閱讀更多關於版本控制的運作方式。


## Outlook 增益集功能

Outlook 增益集提供許多豐富的功能，可以用來支援各種案例。



|**功能**|**說明**|
|:-----|:-----|
|關聯式啟動|Outlook 內容增益集可以根據下列準則啟動︰<ul><li>(預設值) 針對信箱或行事曆中的任何項目</li><li>針對特定項目類型 (電子郵件，會議要求郵件或約會)</li><li>針對項目郵件類別</li><li>關於郵件或約會中的特定實體，請參閱[內容 Outlook 增益集](contextual-outlook-add-ins.md)。</li><li>根據特定規則或規則運算式，請參閱 [Outlook 增益集的啟用規則](manifests/activation-rules.md)和[使用規則運算式的啟用規則來顯示 Outlook 增益集](use-regular-expressions-to-show-an-outlook-add-in.md)。</li><li>關於屬性的字串相符項目，請參閱[使 Outlook 項目中的字串與已知的實體相符](match-strings-in-an-item-as-well-known-entities.md)</li></ul>|
|延伸模組|Outlook 延伸模組會將您的增益集與 Outlook 導覽列整合。如需詳細資訊，請參閱 [將您的 Outlook 增益集與 Outlook 的導覽列整合](../outlook/extension-module-outlook-add-ins.md)。延伸模組僅在 Outlook 2016 for Windows 可供使用。|
|增益集命令|Outlook 增益集命令會提供從功能區初始特定增益集動作的方式。它們只適用於套用至所有電子郵件或事件的延伸模組和增益集。如需詳細資訊，請參閱 [Outlook 的增益集命令](../outlook/add-in-commands-for-outlook.md)。 |
|漫遊設定|Outlook 增益集可以儲存特定於使用者的信箱的資料，您可以在後續的 Outlook 工作階段存取。如需詳細資訊，請參閱[取得和設定 Outlook 增益集的增益集中繼資料](../outlook/metadata-for-an-outlook-add-in.md)。 |
|自訂屬性|Outlook 增益集可以儲存特定於使用者信箱中項目的資料，您可以在後續的 Outlook 工作階段存取。如需詳細資訊，請參閱[取得和設定 Outlook 增益集的增益集中繼資料](../outlook/metadata-for-an-outlook-add-in.md)。|
|取得附件或整個選取的項目|關聯式 Outlook 增益集可以從伺服器端存取附件和整個選取的項目。 請參閱下列主題：<ul><li>附件 - 請參閱 [從伺服器取得 Outlook 項目的附件](get-attachments-of-an-outlook-item.md) 和 [在 Outlook 中新增及移除撰寫格式項目的附件]add-and-remove-attachments-to-an-item-in-a-compose-form.md)</li><li>整個選取的項目 - 這類似於使用回呼權杖以取得附件。 請參閱下列主題：<ul><li>[Office.context.mailbox](../../reference/outlook/Office.context.mailbox.md) 中的 **mailbox.getCallbackTokenAsync**- 提供回乎權杖以識別 Exchange Server 的增益集伺服器端程式碼。</li><li>[Office.context.mailbox](../../reference/outlook/Office.context.mailbox.item.md) 中的 **item.itemId**- 識別使用者讀取的項目，以及伺服器端程式碼取得的項目。</li><li>
  [Office.context.mailbox](../../reference/outlook/Office.context.mailbox.md) 中的 **mailbox.ewsUrl**- 提供 EWS 端點 URL，與回呼權杖和項目 ID 搭配，伺服器端程式碼可以用來存取 [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4(Office.15).aspx) EWS 作業，以取得整個項目。</li></ul></li></ul>|
|使用者設定檔|郵件增益集可以存取顯示名稱、電子郵件地址和使用者設定檔中的時區。如需詳細資訊，請參閱 [UserProfile](../../reference/outlook/Office.context.mailbox.userProfile.md) 物件。|

## 開始建置 Outlook 增益集

若要開始建置 Outlook 增益集，請參閱[開始使用 Office 365 的 Outlook 增益集](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)或[將您的 Outlook 增益集與 Outlook 的導覽列整合](../outlook/extension-module-outlook-add-ins.md)


## 其他資源

如需了解適用於一般開發 Office 增益集的概念，請參閱下列各項︰

- [Office 增益集的設計指導方針](../../docs/design/add-in-design.md)

- [開發 Office 增益集的最佳做法](../../docs/design/add-in-development-best-practices.md)

- [授權您的 Office 和 SharePoint 增益集](http://msdn.microsoft.com/library/3e0e8ff6-66d6-44ff-b0c2-59108ebd9181%28Office.15%29.aspx)

- [將 Office 和 SharePoint 增益集和 Office 365 Web 應用程式提交給 Office 市集](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)

- [JavaScript API for Office](../../reference/javascript-api-for-office.md)

- [Outlook 增益集資訊清單](../outlook/manifests/manifests.md)

