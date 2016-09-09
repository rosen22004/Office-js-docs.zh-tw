
# 疑難排解 Office 增益集的使用者錯誤

使用者有時候可能會在使用您開發的 Office 增益集時發生問題。例如，增益集無法載入或無法存取。使用本文的資訊，協助解決使用者會遇到的 Office 增益集的常見問題。 

您也可以使用 [Fiddler](http://www.telerik.com/fiddler) 來識別及偵錯增益集的問題。

解決使用者問題後，您可以[在 Office 市集中直接回應客戶檢閱](https://msdn.microsoft.com/library/jj635874.aspx)。

## 常見錯誤和疑難排解步驟

下表列出使用者可能會遇到的常見錯誤訊息，和您的使用者可以用來解決錯誤的步驟。



|**錯誤訊息**|**解決方案**|
|:-----|:-----|
|應用程式錯誤︰無法存取目錄|請確認防火牆設定。「目錄」參考至 Office 市集。此訊息表示使用者無法存取 Office 市集。|
|應用程式錯誤︰無法啟動此應用程式。關閉此對話方塊以忽略問題，或按一下 [重新啟動] 再試一次。|請確認已安裝最新的 Office 更新，或下載[Office 2013 更新](https://support.microsoft.com/en-us/kb/2986156/)。|
|錯誤：物件不支援屬性或方法 'defineProperty'|確認 Internet Explorer 不是在相容性模式中執行。 移至 [工具] > [相容性檢視設定]****。|
|很抱歉，我們無法載入應用程式，因為您的瀏覽器版本不受支援。如需支援的瀏覽器版本清單，請按一下這裡。|確定瀏覽器支援 HTML5 本機存放區，或重設您的 Internet Explorer 設定。如需支援瀏覽器的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)|

## Outlook 增益集無法正確運作

如果 Windows 上執行的 Outlook 增益集無法正確運作，請嘗試在 Internet Explorer 中開啟指令碼偵錯。 


- 移至 [工具] > [網際網路選項]**** > **[進階]**。
    
- 在 [瀏覽]**** 下，取消核取 [停用指令碼偵錯 (Internet Explorer)]**** 和 [停用指令碼偵錯 (其他)]****。
    
我們建議您只有在疑難排解問題時才取消核取這些設定。 如果維持取消核取，您就會在瀏覽時看到提示。 解決問題後，重新核取 [停用指令碼偵錯 (Internet Explorer)]**** 和 [停用指令碼偵錯 (其他)]****。


## 增益集不會在 Office 2013 中啟動

如果增益集未在使用者執行下列步驟時啟動︰


1. 使用其在 Office 2013 的 Microsoft 帳戶登入。
    
2. 為其 Microsoft 帳戶啟用兩步驟驗證。
    
3. 當他們嘗試插入增益集出現提示時，確認其身分識別。
    
請確認已安裝最新的 Office 更新，或下載[Office 2013 更新](https://support.microsoft.com/en-us/kb/2986156/)。


## 其他資源



- [在 Office Online 中偵錯增益集](../testing/debug-add-ins-in-office-online.md)
    
- [在 iPad 和 Mac 上側載 Office 增益集](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [在 iPad 和 Mac 上偵錯 Office 增益集](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
- [在 Visual Studio 中建立和偵錯 Office 增益集](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
    
- [部署和安裝 Outlook 增益集以進行測試](../outlook/testing-and-tips.md)
    
