
# <a name="publish-task-pane-and-content-add-ins-to-an-add-in-catalog-on-sharepoint"></a>在 SharePoint 上發佈工作窗格和內容增益集至增益集目錄

增益集目錄是 SharePoint Web 應用程式或 SharePoint Online 租用中的專用網站集合，其裝載 Office 和 SharePoint 增益集的文件庫。系統管理員可以將 Office 增益集資訊清單檔案，上載到其組織內使用的增益集目錄。當系統管理員將增益集目錄註冊為信任的目錄時，使用者可以從 Office 用戶端應用程式中的插入 UI 插入增益集。

>**附註：**SharePoint 的增益集目錄不支援在[增益集資訊清單](../overview/add-in-manifests.md)之 VersionOverrides 節點實作的增益集功能。

Office 2016 for Mac 不支援 SharePoint 目錄。若要將 Office 增益集部署到 Mac 用戶端，您必須將它們提交到 [Office 市集](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)。   

## <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a>若要在 SharePoint 上設定增益集目錄

1. 瀏覽至 [管理中心網站] ([開始] > **[所有程式]** > **[Microsoft SharePoint 2013 產品]** > **[SharePoint 2013 管理中心]**)。
    
2. 在左邊工作窗格中，選擇 [增益集]。
    
3. 在 [增益集] 頁面的 [增益集管理] 下，選擇 [管理增益集目錄]。
    
4. 在 [管理增益集目錄] 頁面上，確定在 [Web 應用程式選取器] 中選取正確的 Web 應用程式。
    
5. 選擇 [檢視網站設定]。
    
6. 在 [網站設定] 頁面上，選擇 [網站集合管理員] 來指定網站集合管理員，然後選擇 [確定]。
    
7. 若要將網站權限授與給使用者，請選擇 [網站權限]，然後選擇 [授與權限限]。
    
8. 在 [共用 '應用程式目錄網站'] 對話方塊中，指定一或多個網站使用者，設定它們的適當權限、選擇性地設定其他選項，然後選擇 [共用]。
    
9. 若要將增益集加入至 Office 增益集的增益集目錄中，請選擇 [Office 增益集]。

## <a name="to-set-up-an-add-in-catalog-on-office-365"></a>若要在 Office 365 上設定增益集目錄

1. 在 Office 365 系統管理中心頁面上，選擇 [管理]，然後選擇 [SharePoint]。
    
2. 在左邊工作窗格中，選擇 [增益集]。
    
3. 在 [增益集] 頁面上，選擇 [增益集目錄]。
    
4. 在 [增益集目錄網站] 頁面上，選擇 [確定] 以接受預設選項，並建立新的增益集目錄網站。
    
5. 在 [建立增益集目錄網站集合] 頁面上，指定增益集目錄網站的標題。
    
6. 指定網站位址。
    
7. 將 [儲存區配額] 設說最小的可能值 (目前為 110)。您只會在此網站集合上安裝增益集套件，它們都很小。
    
8. 將 [伺服器資源配額] 設為 0 (零)。(伺服器資源配額與節流執行不良的沙箱化解決方案有關，但您不會在您的增益集目錄網站上安裝任何沙箱化解決方案。)
    
9. 選擇 [確定]。
    
要將增益集新增至增益集目錄網站，請瀏覽至您剛建立的網站。在左邊的瀏覽窗格中，選擇 [Office 增益集]，然後在資訊清單檔案中上載 Office 增益集，選擇 [新的增益集]。    

## <a name="publish-to-an-add-in-catalog"></a>發佈至增益集目錄


1. 瀏覽至增益集目錄：

    1- 開啟 SharePoint 管理中心主頁面。
    
    2- 選取 [增益集]。
    
    3- 選取 [管理增益集目錄]。
    
    4- 選擇提供的連結，然後選擇左方導覽列的 [Office 增益集]。
    
2. 選擇 [按一下以加入新項目] 連結。
    
3. 選擇 [瀏覽]，然後指定要上傳的[資訊清單](../../docs/overview/add-in-manifests.md)。
    
    [Office 增益集] 對話方塊現在提供此目錄中的內容和工作窗格增益集。若要存取這些增益集，請選擇 [插入] 索引標籤上 [我的增益集]，然後選擇 [我的組織]。
    
將增益集資訊清單上傳至 Office 增益集目錄後，使用者可以執行下列步驟來存取增益集︰


1. 在 Office 應用程式中，移至 [檔案] > **[選項]** > **[信任中心]** > **[信任中心設定]** > **[受信任的增益集目錄]**。
    
2. 指定增益集目錄的_父 SharePoint 網站集合_的 URL。例如，如果 Office 增益集目錄的 URL 為：
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    單獨指定父網站集合的 URL：
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. 關閉並重新開啟 Office 應用程式。增益集目錄將可用於 [Office 增益集] 對話方塊。
    
或者，系統管理員可以使用群組原則，在 SharePoint 上指定 Office 增益集目錄 。如需詳細資訊，請參閱 TechNet 上 [Office 增益集概觀](https://technet.microsoft.com/en-us/library/jj219429.aspx)的<使用群組原則來管理使用者可安裝及使用 Office 增益集的方式> 一節。

