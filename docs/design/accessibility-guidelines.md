# <a name="accessibility-guidelines-for-office-add-ins"></a>Office 增益集的協助工具指導方針

當您設計和開發 Office 增益集時，您會希望確保所有潛在使用者和客戶都能夠順利使用您的增益集。遵照下列指導方針，可以確保所有對象都能存取您的解決方案。

##<a name="design-for-multiple-input-methods"></a>多種輸入方法的設計

- 確保使用者只使用鍵盤也可以執行操作。使用者應該可以使用 Tab 鍵和方向鍵組合，在頁面上所有可操作項目間移動。
- 在行動裝置上，當使用者透過觸控操作控制項，裝置應該提供有用的音訊回饋。
- 為所有互動式控制項提供有幫助的標籤。 

##<a name="make-your-add-in-easy-to-use"></a>將增益集設計成容易使用

- 在 UI 中切勿只使用單一屬性 (例如顏色、大小、形狀、位置、方向或聲音) 來傳遞意義。
- 避免意外的內容變動，例如未經使用者動作即將焦點移動至不同的 UI 項目。
- 提供可以驗證、確認或反向所有繫結動作的方法。
- 提供可以暫停或停止媒體 (例如音訊和視訊) 的方法。
- 不要強制規定使用者動作的時間。

##<a name="make-your-add-in-easy-to-see"></a>將增益集設計成容易查看

- 避免意外的顏色變動。
- 提供有意義且即時的資訊來說明 UI 項目、標題及標頭、輸入項目及錯誤。確認控制項的名稱可充分說明控制項的用途。
- 遵循色彩對比的[標準指導方針](http://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html)。

##<a name="account-for-assistive-technologies"></a>容納輔助技術

- 避免使用會干擾輔助技術的功能，包括視訊、音訊或其他互動方式。
- 不要提供影像格式的文字。螢幕助讀程式無法讀取影像中的文字。
- 提供能讓使用者調整所有音訊來源的音量，或將其設為靜音的方法。
- 提供能讓使用者開啟音訊來源隨附字幕或音訊描述的方法。
- 提供音效以外的提醒使用者方法，例如視覺提示或振動。

##<a name="accessibility-resources"></a>協助工具資源

- [網頁內容協助工具指導方針 (WCAG) 2.0](http://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [套用 WCAG 2.0 至非網頁資訊及通訊技術 (WCAG2ICT) 的指引](http://www.w3.org/TR/wcag2ict/)
- [資訊及通訊技術 (ICT) 的協助工具需求 (歐洲標準)](http://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf)


