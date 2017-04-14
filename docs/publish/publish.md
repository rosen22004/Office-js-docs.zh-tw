
# <a name="deploy-and-publish-your-office-add-in"></a>部署及發佈 Office 增益集

您可以使用下列其中一種方法來部署 Office 增益集，以供測試之用或散發給使用者。

|**方法**|**Use...**|
|:---------|:------------|
|[側載](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|能做為部署程序的一部分，測試增益集在 Windows、Office Online、iPad 或 Mac 上的執行狀況。|
|[集中式部署](centralized-deployment.md)|將增益集散發給雲端或混合部署中貴組織的使用者。|
|[Office Store](https://dev.office.com/officestore/docs/submit-to-the-office-store)|將增益集公開散發給使用者。|
|[SharePoint 目錄](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|將增益集散發給內部部署中貴組織的使用者。|
|[Exchange 伺服器](#outlook-add-in-deployment)|在內部部署或線上環境中，將 Outlook 增益集散發給使用者。|

可用的選項視您鎖定的 Office 主應用程式和建立的增益集類型而定。

>**附註：**如果您打算將增益集發佈至 Office 市集中，請確定您符合 [Office 市集驗證原則](https://msdn.microsoft.com/en-us/library/jj220035.aspx)。例如，若要通過驗證，增益集必須可以在所有支援所定義方法的平台上運作 (如需詳細資料，請參閱 [4.12 節](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably)與 [Office 增益集主應用程式與可用性頁面](https://dev.office.com/add-in-availability))。

如需使用者如何取得、插入及執行增益集的相關資訊，請參閱[開始試用 Office 增益集](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)。


## <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Word、Excel 及 PowerPoint 增益集的部署選項

| 擴充點            | 側載 | Office 365 系統管理中心 (預覽) |Office 市集| SharePoint 目錄*  |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| 內容         | X           | X                  | X                               | X|
| 工作窗格       | X           | X                  | X                               | X|
| 命令           | X           | X                  | X                               |  |

&#42; SharePoint 目錄不支援 Office 2016 for Mac。

## <a name="deployment-options-for-outlook-add-ins"></a>Outlook 增益集的部署選項

| 擴充點     | 側載 | Exchange Server | Office 市集 |
|:---------|:-----------:|:---------------:|:------------:|
| 郵件應用程式 | X           | X               | X            |
| 命令  | X           | X               | X            |


## <a name="centralized-deployment-via-the-office-365-admin-center-preview"></a>透過 Office 365 系統管理中心 (預覽) 的集中式部署

管理使用者可透過 Office 365 系統管理中心，輕鬆將 Word、Excel 及 PowerPoint 增益集部署至其組織內的使用者或群組。使用者可立即在 Office 應用程式中使用透過系統管理中心部署的增益集，而無需進行用戶端設定。您可以透過集中式部署，部署內部增益集以及 ISV 提供的增益集。

如需詳細資訊，請參閱[使用集中式部署發佈增益集](centralized-deployment.md)。


## <a name="sharepoint-catalog-deployment"></a>SharePoint 目錄部署

SharePoint 增益集目錄是一特殊網站的集合，您可建立用來裝載 Word、Excel 及 PowerPoint 增益集。因為 SharePoint 目錄不支援資訊清單 [VersionOverrides] 節點中實作的新增益集功能 (包含增益集命令)，我們建議您透過系統管理中心 (預覽) 使用集中式的部署 (如果可能的話)。依預設，會在工作窗格中開啟透過 SharePoint 目錄部署的增益集命令。

如果您要在內部部署環境中部署增益集，請使用 SharePoint 目錄。如需詳細資訊，請參閱[將工作窗格和內容增益集發佈至 SharePoint 類別目錄](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。

> **附註：**SharePoint 類別目錄不支援 Office 2016 for Mac。若要將 Office 增益集部署到 Mac 用戶端，您必須將它們提交到 [Office 市集]。 

## <a name="outlook-add-in-deployment"></a>Outlook 增益集部署

對於不使用 Azure AD 識別服務的內部部署與線上環境，您可以透過 Exchange 伺服器部署 Outlook 增益集。 

Outlook 增益集部署必要條件：

- Office 365、Exchange Online，或 Exchange Server 2013 或更新版本
- Outlook 2013 或更新版本

若要將增益集指派給租用戶，您可以使用 Exchange 系統管理中心，透過檔案或 URL 來直接上載資訊清單，或透過 Office 市集新增增益集。若要將增益集指派給個別使用者，您必須使用 Exchange PowerShell。如需詳細資訊，請參閱 TechNet 上的[為您的組織安裝或移除 Outlook 增益集](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx)。


## <a name="additional-resources"></a>其他資源

- [部署和安裝 Outlook 增益集以進行測試](../outlook/testing-and-tips.md) 
- [提交至 Office 市集][Office 市集]
- [Office 增益集的設計指導方針](../design/add-in-design)
- [建立有效的 Office 市集增益集](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [疑難排解 Office 增益集的使用者錯誤](../testing/testing-and-troubleshooting.md)

[Office 市集]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
 