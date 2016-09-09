# 使用執行階段記錄來偵錯增益集命令

Office 16 桌面用戶端有一項新功能可記錄有用的資訊。 除了其他功能之外，此工具可協助您診斷增益集資訊清單中的錯誤，尤其是當您建立的是包含增益集命令的資訊清單，會特別有用。 

此功能的完整文件即將推出，同時請參考以下資訊，了解如何使用此功能於剖析包含增集益命令的資訊清單時進行偵錯。

##開啟執行階段記錄

**重要事項**：執行階段記錄會**影響效能**。 僅當您需要偵錯增益集問題時再開啟此功能。

1. 請確定您有支援執行階段記錄的組建。 您需要 **Office 16 桌面**用戶端組建 **16.0.7019** 或更新版本
2. 在 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\` 下方新增 `RuntimeLogging` 登碌機碼 
3. 將機碼的預設值設定為要寫入記錄之檔案的完整路徑。 請參閱[登錄機碼的範例](RuntimeLogging/EnableRuntimeLogging.zip) (解壓縮)

您的登錄看起來應該像這樣︰![](http://i.imgur.com/Sa9TyI6.png)

若要關閉此功能，只需從登錄移除機碼。 

##診斷命令的問題
執行階段記錄有助於偵測不易辨識的**資訊清單問題**，例如資源的識別碼不相符、長度無效等 XSD 結構描述驗證無法偵測的問題。 

可以嘗試以下步驟來進行排解：
 
1. 請遵循[讀我檔案](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/README.md)上的指示來側載您的增益集。 
2. 如果您看不到功能區按鈕專案，且增益集對話方塊沒有出現任何項目，請檢查記錄檔
3. 搜尋在資訊清單中定義的增益集識別碼，以尋找該增益集的所屬訊息。 記錄檔會將此識別碼報告為 `SolutionId` 建議您一次只側載一個增益集，以避免無法辨識出特定增益集的所屬訊息。 

在下面的範例中，RuntimeLogging 協助識別指向不存在資源檔的控制項。 修正方法是更正打錯的字 (如果有的話)，或確實新增該遺失資源。

![](http://i.imgur.com/f8bouLA.png) 

##記錄的已知問題
執行階段記錄仍有已知的錯誤。 您可能會看到數個令人困惑或不當分類的訊息。 例如：

- 系統將 `Unexpected Parsed manifest targeting different host` 之前的 `Medium  Current host not in add-in's host list` 訊息遭不當分類。 這些不是錯誤，您可放心地忽略這些錯誤。
- 訊息 `Unexpected   Add-in is missing required manifest fields  DisplayName` 不包含有問題增益集的 SolutionId。 不過，這很可能與您正在偵錯的增益集無關。 
- 從系統的觀點看來，所有 `Monitorable` 訊息都應該是錯誤。 有時候，它們會指出資訊清單發生問題 (如系統略過拼字錯誤但未造成資訊清單失敗的元素)。 

