
# <a name="sideload-office-add-ins-for-testing"></a>側載 Office 增益集來進行測試

您可以在執行於 Windows 上的 Office 用戶端中安裝 Office 增益集以供測試，方式為使用共用資料夾目錄來發佈資訊清單到網路檔案共用。 

>**附註：**若要在 Office Online 中測試 Office 增益集，請參閱[在 Office Online 中側載 Office 增益集來進行測試](sideload-office-add-ins-for-testing.md)。若要在 iPad 或 Mac 上測試增益集，請參閱[在 iPad 和 Mac 上側載 Office 增益集來進行測試](sideload-an-office-add-in-on-ipad-and-mac.md )。若要測試 Outlook 增益集，請參閱[側載 Outlook 增益集來進行測試](sideload-outlook-add-ins-for-testing.md )。

僅部署資訊清單檔案至共用資料夾目錄。將 Web 應用程式本身部署到網頁伺服器，並在資訊清單檔的 **SourceLocation** 元素中指定 URL。

 >**重要：**為了協助增益集更安全的存取外部資料和服務，增益集應使用安全的通訊協定 (例如超文字傳輸通訊協定安全性 (HTTPS)) 以連接至外部資料和服務。如果增益集使用增益集命令，您必須使用 HTTPS。

## <a name="share-a-folder"></a>共用資料夾

1. 在您要託管增益集的 Windows 電腦上，移至您要使用做為共用資料夾目錄的上層資料夾或磁碟機代號。

2. 開啟資料夾的內容功能表 (按一下滑鼠右鍵)，然後選擇 [屬性]。

3. 開啟 [共用] 索引標籤。

4. 在 [選擇人員...] 頁面上，新增自己以及您想要與之共用增益集的任何人。如果他們都是安全性群組的成員，您可以新增群組。您至少需要資料夾的**讀取/寫入**權限。 

5. 選擇 [共用] >  [完成] >  [關閉]。

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>指定共用資料夾做為受信任目錄

      
3. 在 Excel、Word 或 PowerPoint 中開啟新文件。
    
4. 選擇 [檔案] 索引標籤，然後選擇 [選項]。
    
5. 選擇 [信任中心]，然後選擇 [信任中心設定] 按鈕。
    
6. 選擇 [受信任的增益集目錄]。
    
7. 在 [目錄 URL] 方塊中，輸入共用資料夾目錄的完整網路路徑，然後選擇 [新增目錄]。
    
8. 選取 [顯示於功能表中] 核取方塊，然後選擇 [確定]。

9. 關閉 Office 應用程式，如此您的變更才會生效。
    
## <a name="sideload-your-add-in"></a>側載增益集


1. 將您測試的任何增益集的資訊清單檔案放置到共用資料夾目錄中。

2. 在 Excel、Word 或 PowerPoint 中，選取功能區的 [插入] 索引標籤上的 [我的增益集]。

3. 在 [Office 增益集] 對話方塊頂端，選擇 [共用資料夾]。

4. 選取增益集的名稱，然後選擇 [確定] 以插入增益集。


## <a name="additional-resources"></a>其他資源

- [使用執行階段記錄來偵錯您的資訊清單](../develop/use-runtime-logging-to-debug-manifest.md)
- [發佈 Office 增益集](../publish/publish.md)
    
