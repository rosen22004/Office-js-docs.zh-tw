# <a name="use-runtime-logging-to-debug-the-manifest-for-your-office-add-in"></a>使用執行階段記錄來偵錯 Office 增益集的資訊清單

您可以使用執行階段記錄來偵錯增益集的資訊清單。這項功能可協助您找出及修正 XSD 結構描述驗證未偵測到的資訊清單問題，如資源識別碼不符。執行階段記錄特別適合用來偵錯可實作增益集命令的增益集。  

>**附註︰**執行階段記錄功能目前適用於 Office 2016 桌面。

## <a name="turn-on-runtime-logging"></a>開啟執行階段記錄

>**重要事項**：執行階段記錄會影響效能。僅當您需要偵錯增益集資訊清單問題時再開啟。

1. 確認您執行的是 Office 2016 桌面組建 **16.0.7019** 或更新版本。 
2. 將 `RuntimeLogging` 登錄機碼新增到 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\' 下方。 
3. 將機碼的預設值設定為要寫入記錄之檔案的完整路徑。如需範例，請參閱 [EnableRuntimeLogging.zip](RuntimeLogging/EnableRuntimeLogging.zip)。 

 > **附註︰**將在其中寫入記錄檔的目錄必須已經存在，而且您必須具有寫入權限。 
 
下圖展示登錄的外觀。![登錄編輯器和 RuntimeLogging 登錄機碼的螢幕擷取畫面](http://i.imgur.com/Sa9TyI6.png)

若要關閉此功能，請從登錄移除 `RuntimeLogging` 機碼。 

## <a name="troubleshoot-issues-with-your-manifest"></a>疑難排解資訊清單的問題

若要使用執行階段記錄來疑難排解增益集載入問題︰
 
1. [側載增益集](../testing/sideload-office-add-ins-for-testing.md)以進行測試。 

    >附註：我們建議您只側載要測試的增益集，以便減少記錄檔中的訊息數目。
2. 如果沒有出現任何反應且您未看見增益集 (而且未出現在增益集對話方塊中)，請開啟記錄檔。
3. 在記錄檔內搜尋於資訊清單中定義的增益集識別碼。在記錄檔中，該識別碼會標示為 `SolutionId`。 

在下列範例中，記錄檔識別出指向不存在之資源檔的控制項。對於該範例，修正方法是更正資訊清單中的錯字，或新增遺失的資源。

![指出找不到資源識別碼之項目的記錄檔螢幕擷取畫面](http://i.imgur.com/f8bouLA.png) 

##<a name="known-issues-with-runtime-logging"></a>執行階段記錄的已知問題
在記錄檔中，您可能會看到混淆不清或分類錯誤的訊息。例如：

- 系統將 `Unexpected Parsed manifest targeting different host` 之前的 `Medium   Current host not in add-in's host list` 訊息不當分類為錯誤。
- 如果您看到 `Unexpected    Add-in is missing required manifest fields  DisplayName` 訊息且該訊息不含 SolutionId，該項錯誤很有可能與您正在偵錯的增益集無關。 
- 從系統的觀點看來，所有 `Monitorable` 訊息都應該是錯誤。有時候，它們會指出資訊清單發生問題，如系統略過拼字錯誤但未造成資訊清單失敗的元素。 

##<a name="additional-resources"></a>其他資源

- [側載 Office 增益集來進行測試](../testing/sideload-office-add-ins-for-testing.md)
- [偵錯 Office 增益集](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)