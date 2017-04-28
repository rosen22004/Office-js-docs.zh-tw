# <a name="use-centralized-deployment-to-publish-office-add-ins"></a>使用集中式部署發佈 Office 增益集

管理使用者可透過 Office 365 系統管理中心，輕鬆將 Word、Excel 及 PowerPoint 增益集部署至其組織內的使用者或群組。使用者可立即在 Office 應用程式中使用透過系統管理中心部署的增益集，而無需進行用戶端設定。您可以透過集中式部署，部署內部增益集以及 ISV 提供的增益集。

系統管理中心目前支援下列案例：

- 將新增及更新增益集集中部署至個人、群組或組織。
- 部署至多重平台，包括 Windows 和 Office Online，與即將推出的 Mac。
- 英文語言和全球租用戶的部署。
- 雲端託管增益集部署。
- Office 應用程式啟動時的自動安裝。
- 防火牆內裝載的增益集 URL。
- Office 市集增益集的部署 (即將推出)。

<!--
The admin center also includes a pre-deployment validation checking service.
-->

增益集部署案例的未來投資會著重在 Office 365 系統管理中心。如果您的組織符合所有必要條件，我們建議您透過系統管理中心，將增益集部署至您的組織。

## <a name="prerequisites-for-centralized-deployment"></a>集中式部署的必要條件 

如果您的組織符合下列條件，您可以透過系統管理中心部署增益集︰

- 使用者執行的是 Office 2016 ProPlus 的版本︰
    - Windows 組建 16.0.8027 或更新版本
    - Mac 組建 15.33.170327 或更新版本
- 使用者透過公司或學校帳戶登入 Office 2016。
- 您的組織使用的是 Azure Active Directory (Azure AD) 身分識別服務。
- 使用者的 Exchange 信箱[已啟用 OAuth](https://msdn.microsoft.com/en-us/library/office/dn626019(v=exchg.150).aspx#Anchor_0)。

目前支援下列 Office 用戶端的增益集。 

| Office 應用程式    | Office 2016 for Windows   | Office Online | Office 2016 for Mac   |
|:----------------------|:-------------------------:|:-------------:|:---------------------:|
| Word                  | X                         | X             | X                     |
| Excel                 | X                         | X             | X                     |
| PowerPoint            | X                         | X             | X                     |
| Outlook               | 即將推出               | 即將推出   | 即將推出           |

系統管理中心不支援下列項目：

- Office 2013 (Word、Excel、PowerPoint 或 Outlook)。
- Office for iPad
- SharePoint 增益集。
- 以 COM/VSTO 為基礎的增益集。
- Office Online Server。
- 內部部署目錄服務。

若要部署 SharePoint 增益集或以 Office 2013 為目標的增益集，使用 [SharePoint 增益集目錄](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。

>**重要！**SharePoint 增益集目錄不支援在增益集資訊清單之 [VersionOverrides](../../reference/manifest/versionoverrides.md) 節點實作的增益集功能，例如[增益集命令](../design/add-in-commands.md)。 

若要部署 COM/VSTO 增益集，使用 ClickOnce 或 Windows Installer。如需詳細資訊，請參閱[部署 Office 解決方案](https://msdn.microsoft.com/en-us/library/bb386179.aspx)。

<!-- Need URL on SOC site.
For more information about requirements, see [centralized deployment eligibility]().
-->

## <a name="publish-an-add-in-via-centralized-deployment"></a>透過集中式部署發佈增益集

若要透過集中式部署發佈增益集：

1.    請驗證組織是否符合[集中式部署的必要條件](#prerequisites-for-centralized-deployment)。
2.    在 Office 365 系統管理中心頁面上，選擇 [設定]**** > [服務與增益集]****。
3.    選擇頁面頂端的 [新增 Office 增益集]****。您有下列選項：

    - 從 Office 市集新增增益集。
    - 選擇 [瀏覽]**** 以找出資訊清單 (.xml) 檔案位置。
    - 在提供的欄位中為您的資訊清單輸入 URL。

5.    選擇 [下一步]****。
6.    如果您正從 Office 市集新增增益集，選取該增益集。增益集現在已啟用。 
7.    選擇 [編輯]**** 以將增益集指派給使用者。 
8.    搜尋您想要部署增益集的人員或群組，並選擇其名稱旁的 [新增]****。
9.    選擇 [儲存]****，檢閱增益集設定，然後選擇 [關閉]****。


如果增益集支援增益集命令，針對所有部署增益集的使用者，命令將會顯示在 Office 應用程式的功能區上。 

如果增益集不支援增益集命令，使用者可從 [我的增益集]**** 按鈕進行新增，執行下列動作︰

1.    在 Word 2016、Excel 2016 或 PowerPoint 2016 中，選擇 [插入]**** > [我的增益集]****。
2.    選擇增益集視窗中的 [系統管理員管理]**** 索引標籤。
3.    選擇增益集，然後選擇 [新增]****。 

