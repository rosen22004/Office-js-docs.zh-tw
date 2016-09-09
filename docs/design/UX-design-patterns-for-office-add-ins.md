# Office 增益集的 UX 設計模式。 

設計 Office 增益集時，增益集的 UX 設計應提供可擴充 Office 的絕佳體驗。若要建立出色的增益集，增益集應該提供初次執行體驗、最出色 UX 經驗，以及能夠流暢地在其他頁面間轉換。提供簡潔、最新的 UX 體驗能夠提升使用者忠誠度並讓更多人採用您的增益集。本文將介紹適用於設計人員和開發人員的 UX 資源︰

* 描述根據最佳作法的一般 UX 設計模式。
* 實作 Office Fabric 元件與樣式。
* 實作看起來像預設 Office UI 自然擴充的增益集。 

## 如何開始使用 Office 增益集設計範例資源？

沒有任何先決條件，即可使用這些設計或程式碼資產。若要開始為您的增益集建立絕佳的 UX︰

* 檢閱 UX 設計模式，並找出哪些對增益集來說很重要。例如，挑選其中一個初次執行體驗。
* 執行下列一項或多項操作：
	* 將程式碼檔案複製到您的增益集專案，並開始自訂以符合您的需求。您將會需要 [common.js 檔案](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/)、[資產資料夾](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets)，和您所需之設計模式的程式碼資料夾。請參閱下列連結。
	* 建立您自己的 UX 設計時，請下載參考 PDF 並將它們作為指南使用。請參閱下列連結。
	* 下載 Adobe Illustrator 檔案並加以編輯，來模擬您自己的增益集設計。可在[這裡](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Source%20Files)取得檔案。
 

## 初次執行

初次執行體驗是使用者第一次開啟增益集時的體驗。下面列出您可在增益集中包含的初次執行設計模式。其中每個影像如下。

* **啟動步驟**為使用者提供步驟排序清單，即可開始使用增益集。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_StepsToStart.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/instruction-step))
* **值**傳達了增益集的價值主張。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_ValuePlacemat.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/value-placemat))
* **影片**在使用者開始增益集前會向其顯示影片。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_VideoPlacemat.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/video-placemat))
* **逐步解說**在使用者開始使用增益集前，會先帶他們了解一系列功能或資訊。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_PagingPanel.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/walkthrough))
* [Office 市集](https://msdn.microsoft.com/zh-tw/library/office/jj220033.aspx)的系統可提供使用者增益集的試用版，但是如果您想要完全控制試用版的 UI 經驗，請使用下列範本︰
	* **試用版**向使用者示範如何開始使用增益集試用版。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_TrialVersion.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat))
	* **試用版功能**會告知使用者，他們嘗試使用的功能在增益集試用版中無法使用 ([程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat-feature))


> 附註：請考慮對您的案例來說，向使用者顯示一次或多次初次執行體驗是否很重要。例如，如果使用者是定期使用您的增益集，它們可能會忘記如何使用。再次看到初次執行體驗對這些使用者來說可能很有幫助。 

 <table>
 <tr><th>開始步驟</th><th>值</th><th>影片</th></tr>
 <tr><td>![instruction steps" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/instruction.step.PNG)</td><td>![value placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/value.placemat.PNG)</td><td>![video placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/video.placemat.PNG)</td></tr>
 </table>

 <table>
 <tr><th>逐步解說第一頁</th><th>試用版</th><th>試用版功能</th></tr>
 <tr><td>![walkthrough 1" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/walkthrough1.PNG)</td><td>![trial placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.PNG)</td><td>![trial placemat feature" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.feature.PNG)</td></tr>
 </table> 


## 通用和品牌推廣

* **登陸頁面**是初次執行體驗或登入程序後使用者第一個瀏覽區域。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Helpful%20Templates/AddIn_Template_Standard_Layout.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/generic/landing-page))

<table>
 <tr><th>登陸</th></tr>
 <tr><td>![landing page" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/landing.page.PNG)</td></tr>
 </table>

## 通知

增益集有多種方式，可以通知使用者事件 (例如錯誤) 或進度。下表列出這些技術。其中每個影像如下。

* **內嵌對話方塊**會在工作窗格顯示對話方塊，其中提供資訊，並使用按鈕或其他控制項選擇性地提供互動式體驗。請考慮使用其中一種來提示使用者確認動作。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Embedded_Dialog.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/embedded-dialog))
* **內嵌訊息**表示錯誤、成功或資訊，而且它可以出現在工作窗格中指定的位置。例如，如果使用者在文字方塊中輸入格式不正確的電子郵件地址，錯誤訊息就會出現在文字方塊的正下方。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_Inline_Message.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/inline-message))
* **訊息橫幅**提供資訊，並選擇性地提供簡單的動作呼叫，該橫幅可摺疊成一行、展開成好幾行，或將其關閉。請考慮在增益集啟動時，使用訊息橫幅來報告服務更新或有用的提示。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_messagebanner.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/message-banner))
* **進度列**指出長時間執行、同步處理程序 (例如使用者在能夠採取任何動作前必須完成的設定工作) 的進度。它是個別的插入式頁面，其中同時會強調增益集的品牌。當程序可以定期將程序還需多久時間的量值傳回增益集時，請使用進度列。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/progress-bar))
* **載入狀態圓環**表示長時間執行、同步處理程序正在進行中，但不會顯示還需多久的時間。它是個別的插入式頁面，其中同時會強調增益集的品牌。當增益集無法知道處理程序還需多久時，請使用載入狀態圓環。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/spinner))
* **快顯通知**提供幾秒鐘後即會淡出的簡短訊息。因為使用者可能不會看到訊息，只對非必要的資訊才使用快顯通知。這是一個好選擇，可用在遠端系統中通知使用者接收電子郵件之類的事件。([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_toast.pdf "PDF")、[程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/toast))

 <table>
 <tr><th>內嵌的對話方塊</th><th>內嵌訊息</th><th>訊息橫幅</th></tr>
 <tr><td>![embedded dialog" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/embedded.dialog.PNG)</td><td>![inline message" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/inline.message.PNG)</td><td>![message banner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/message.banner.PNG)</td></tr>
 </table>

 <table>
 <tr><th>進度列</th><th>載入狀態圓環</th><th>快顯通知</th></tr>
 <tr><td>![progress bar" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/progress.bar.PNG)</td><td>![spinner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/spinner.PNG)</td><td>![toast" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/toast.PNG)</td></tr>
 </table>

## 已知問題

* 執行增益集專案以外的某些程式碼檔案會擲回 JavaScript 錯誤。 
	* 解決方案：請確定您在 Office 增益集專案中加入這些檔案。 
	
## 其他資源

* [開發 Office 增益集的最佳作法](https://dev.office.com/docs/add-ins/design/add-in-development-best-practices)
* [Office UI Fabric](http://dev.office.com/fabric/)
