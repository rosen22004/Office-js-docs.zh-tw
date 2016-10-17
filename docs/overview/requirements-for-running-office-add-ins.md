
# <a name="requirements-for-running-office-add-ins"></a>執行 Office 增益集的需求


本文說明執行 Office 增益集的軟體及裝置需求。

>**附註︰**如需目前 Office 增益集受支援所在的高階檢視，請參閱 [Office 增益集主應用程式和平台可用性](http://dev.office.com/add-in-availability)頁面。 


## <a name="server-requirements"></a>伺服器需求

若要安裝和執行任何 Office 增益集，您必須先將增益集的 UI 和程式碼的資訊清單和網頁檔案，部署至適當的伺服器位置。

對所有類型的增益集 (內容、Outlook，和工作窗格增益集和增益集命令)，您需要將增益集的網頁檔案部署到 Web 伺服器或 Web 裝載服務，例如 [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md)。


 >**附註：**在 Visual Studio 中開發和偵錯增益集時，Visual Studio 會在本機使用 IIS Express 部署和執行增益集的網頁檔案，而且您不需要其他 Web 伺服器。同樣地，當您在瀏覽器中使用 Napa 開發及偵錯時，它會從您用來登入 Napa 的帳戶所相關的儲存體來部署並執行增益集的網頁檔案。

對於內容和工作窗格增益集，在支援的 Office 主機應用程式 (Access Web 應用程式、Word、Excel、PowerPoint 或 Project) 中，您在 SharePoint 上還需要[增益集目錄](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)以上載增益集的 XML 資訊清單檔案。

若要測試和執行 Outlook 增益集，使用者的 Outlook 電子郵件帳戶必須位於 Exchange 2013 或更新版本，可透過 Office 365，Exchange Online 或內部部署安裝取得。使用者或系統管理員會在該伺服器上安裝 Outlook 增益集的資訊清單檔案。

 >**附註：** Outlook 中的 POP 和 IMAP 電子郵件帳號不支援 Office 增益集。




## <a name="client-requirements:-windows-desktop-and-tablet"></a>用戶端需求Windows 桌上型電腦和平板電腦

需要下列軟體，才能為 Windows 架構桌上型電腦、膝上型電腦或平板電腦裝置上執行的支援 Office 桌面用戶端或 Web 用戶端開發 Office 增益集︰


- 對於 Windows x86 和 x64 桌上型電腦，及 Surface Pro 之類的平板電腦︰

    - 32 位元或 64 位元版本的 Office 2013 或更新版本，在 Windows 7 或更新版本上執行。

    - 如果您正在為其中一個 Office 桌面用戶端測試或執行 Office 增益集，則為 Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013 或較新版本的 Office 用戶端。Office 桌面用戶端可以安裝在內部，或在用戶端電腦上透過按一下執行安裝。

- Internet Explorer 9 或更新版本，必須先安裝，但不一定是預設瀏覽器。若要支援 Office 增益集，做為主機的 Office 用戶端會使用 Internet Explorer 9 或更新版本的瀏覽器元件。

- 預設瀏覽器為下列其中一項︰Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13 或其中一個瀏覽器的更新版本。

- 記事本、[Visual Studio 和 Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs)，或協力廠商 Web 開發工具之類的 HTML 和 JavaScript 編輯器。


## <a name="client-requirements:-os-x-desktop"></a>用戶端需求OS X 桌上型電腦

作為 Office 365 一部分散發的 Outlook for Mac 支援 Outlook 增益集。在 Outlook for Mac 上執行 Outlook 增益集，與 Outlook for Mac 本身具有相同的需求︰必須至少是作業系統 OS X v10.10 "Yosemite"。由於 Outlook for Mac 會將 WebKit 做為配置引擎來呈現增益集頁面，但也沒有任何其他的瀏覽器相依性。

以下是支援 Office 增益集之 Office for Mac 的最低用戶端版本︰
- Word for Mac 15.18 版 (160109) 
- Excel for Mac 15.19 版 (160206) 
- PowerPoint for Mac 15.24 版 (160614)

## <a name="client-requirements:-browser-support-for-office-online-web-clients-and-sharepoint"></a>用戶端需求Office Online Web 用戶端和 SharePoint 的瀏覽器支援

支援 ECMAScript 5.1、HTML5 和 CSS3 的任何瀏覽器，例如 Internet Explorer 9、Chrome 13、Firefox 5、Safari 5.0.6 或較新版的這些瀏覽器。


## <a name="client-requirements:-non-windows-smartphone-and-tablet"></a>用戶端需求︰非 Windows 智慧型手機和平板電腦

特別針對裝置的 OWA，以及智慧型手機和非 Windows 平板電腦裝置瀏覽器中執行的 Outlook Web App，需要下列軟體才能測試和執行 Outlook 增益集。


| 主機應用程式 | 裝置 | 作業系統 | Exchange 帳號 | 行動瀏覽器 |
|:-----|:-----|:-----|:-----|:-----|
|OWA for Android|Android 智慧型手機。[Android 的 OS](https://developer.android.com/guide/practices/screens_support.html) 在技術上將這些裝置視為「小型」或「一般」。|Android 4.4 KitKat 或更新版本|在商務用 Office 365 或 Exchange Online 的最新更新|Android 的原生增益集，不適用瀏覽器|
|OWA for iPad|iPad 2 或更新版本|iOS 6 或更新版本|在商務用 Office 365 或 Exchange Online 的最新更新|iOS 的原生增益集，不適用瀏覽器|
|OWA for iPhone|iPhone 4s 或更新版本|iOS 6 或更新版本|在商務用 Office 365 或 Exchange Online 的最新更新|iOS 的原生增益集，不適用瀏覽器|
|Outlook Web App|iPhone 4 或更新版本、iPad 2 或更新版本、iPod Touch 4 或更新版本|iOS 5 或更新版本|在 Office 365、Exchange Online，或在 Exchange Server 2013 或更新版本上的內部部署|Safari|


## <a name="additional-resources"></a>其他資源

- [Office 增益集平台概觀](../../docs/overview/office-add-ins.md)
- [Office 增益集主應用程式和平台可用性](http://dev.office.com/add-in-availability)

