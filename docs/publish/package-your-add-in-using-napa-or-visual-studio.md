
# <a name="package-your-add-in-using-napa-or-visual-studio-to-prepare-for-publishing"></a>使用 Napa 或 Visual Studio 封裝增益集以準備發佈

您的 Office 增益集套件包含您將用來發佈增益集的 XML 檔案。您必須分別發佈專案的 Web 應用程式檔案。

## <a name="package-an-office-add-in-that-you-create-by-using-napa"></a>封裝使用 Napa 所建立的 Office 增益集



1. 在 Napa 的頁面側邊，選擇 [發佈] 按鈕 (![發佈按鈕](../../images/Apps_NAPA_Publish.png))
    
2. 在 [發佈設定] 對話方塊方塊中，選擇 [下一步]。
    
3. 提供將裝載增益集的內容檔案 (例如，專案的預設 HTML 和 JavaScript 檔案) 的網站 URL，然後再選擇 [發佈]。
    
4. 在[發佈成功] 對話方塊中，選擇 [發佈位置] 連結。
    
    文件庫會出現，其中包含增益集的 XML 資訊清單檔和 Web 內容檔。 
    
接下來，手動將樣式表、JavaScript 檔案和 HTML 檔案的 Web 內容檔案，複製到主控您在 [發佈設定] 對話方塊中提供之網站的 Web 伺服器。

您現在可以上載 XML 資訊清單至適當的位置，以[發佈增益集](../publish/publish.md)。 


## <a name="deploy-your-web-project-and-package-your-add-in-by-using-visual-studio-2015"></a>部署您的 Web 專案，並使用 Visual Studio 2015 封裝增益集



### <a name="to-deploy-your-web-project"></a>部署 Web 專案


1. 在 [方案總管] 中，開啟增益集專案的快顯功能表，然後選擇 [發佈]。
    
    **發佈增益集**頁面隨即出現。
    
2. 在 [目前的設定檔] 下拉式清單中，選取設定檔或選擇 [新增...] 來建立新設定檔。
    
     >**附註**  發佈設定檔會指定您要部署的目的地伺服器、登入伺服器的憑證、要部署的資料庫，以及其他部署選項。

    如果您選擇 [新增...]，[建立發佈設定檔] 精靈隨即出現。您可以使用這個精靈來從如 Microsoft Azure 的網站主控提供者匯入發佈設定檔，或在下一個程序中建立新的設定檔、新增您的伺服器、憑證和其他設定。
    
    如需有關匯入發佈設定檔，或建立新的發佈設定檔的詳細資訊，請參閱[建立發佈設定檔](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile)。
    
3. 在 [發佈增益集] 頁面上，選擇 [部署 Web 專案] 連結。
    
    [發佈 Web] 對話方塊隨即顯示。如需使用此精靈的詳細資訊，請參閱[操作方法：使用 Visual Studio 中單鍵發行來部署 Web 專案](http://msdn.microsoft.com/en-us/library/dd465337.aspx)。
    

### <a name="to-package-your-add-in"></a>封裝增益集


1. 在 [發佈增益集] 頁面上，選擇 [封裝增益集] 連結。
    
    便會顯示 [發佈 Office 和 SharePoint 增益集] 精靈。
    
2. 在 [您的網站架設在哪裡?] 下拉式清單中，選取或輸入將裝載增益集內容檔的網站 URL，然後選擇 [完成]。
    
    您必須指定以 HTTPS 前置詞開頭的位址，才能完成這個精靈。一般情況下，最好的方法是讓您的網站使用 HTTPS 端點，但如果您不打算將增益集發佈至 Office 市集則不需要這麼做。建立套件之後，您可以在記事本中開啟資訊清單，並將 HTTPS 前置詞取代為 HTTP 前置詞。如需詳細資訊，請參閱[為什麼我的增益集必須受到 SSL 安全保護？](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7)。 
    
     >**附註**  Azure 的網站會自動提供 HTTPS 端點。

    Visual Studio 會產生檔案，您需要該檔案以發佈您的增益集，然後開啟發佈輸出資料夾。 
    
如果您打算送出增益集至 Office 市集，您可以選擇 [執行驗證檢查] 連結來識別會使您的增益集無法被接受的問題。您應該先解決所有問題，再將增益集提交至市集 。

您現在可以上載 XML 資訊清單至適當的位置，以[發佈增益集](../publish/publish.md)。您會在 `OfficeAppManifests` 資料夾的 `app.publish` 中找到 XML 資訊清單。例如：

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>其他資源



- [發佈 Office 增益集](../publish/publish.md)
    
-   [將 Office 和 SharePoint 增益集和 Office 365 Web 應用程式提交給 Office 市集](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
