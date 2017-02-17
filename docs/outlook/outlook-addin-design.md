# <a name="outlook-add-in-design-guidelines"></a>Outlook 增益集設計指導方針

對協力廠商而言，增益集是讓 Outlook 功能超越我們的核心功能集的絕佳方法。增益集讓使用者不需離開其收件匣，即可存取協力廠商的經驗、工作和內容。Outlook 增益集一旦安裝，即可在每個平台和裝置上使用。下列高階指導方針將協助您設計和建置具吸引力的增益集，其可將您應用程式的最佳功能放入 Outlook 中 – 適用於 Windows、Web、iOS、Mac 和 Android (即將推出)。

## <a name="principles"></a>原則

1. **著重於一些重要工作；妥善處理這些工作**

    設計最理想的增益集具有易於使用、聚焦的特色，以及提供實際值給使用者。因為增益集將在 Outlook 內部執行，所以會額外強調這個原則。Outlook 是生產力應用程式，使用者可以在此處理事情。

    您將會是我們的經驗延伸，因此務必確保您所啟用的案例感覺像是 Outlook 中渾然天成的案例。請仔細思考您常用的哪個使用案例，可以從我們的電子郵件和行事曆經驗獲得最大益處。

    增益集不應嘗試執行您的應用程式所做的一切。焦點應該放在 Outlook 內容中最常使用的相稱動作。請思考您的行動呼籲 (Call To Action) 並清楚說明當工作窗格開啟時使用者應該怎麼做。

2. **儘可能保持原生增益集**

    應使用 Outlook 執行所在平台的原生模式來設計增益集。若要達成此目標，請務必遵守及實作每個平台所提出的互動和視覺效果指導方針。Outlook 有其自己的指導方針，也一定要將其列入考量。設計完善的增益集會適當地融合您的經驗、平台及 Outlook。

    這表示您的增益集在 iOS 版 Outlook 和 Android 版 Outlook 上執行時的視覺化方式會不同 (當我們推行該支援時)。我們建議您採用 [Framework7](https://framework7.io/)，以協助您使用樣式設定。我們將會公佈更新的指導分針，尤其是針對 Android，因為我們未來會啟動 Android 版 Outlook 的增益集支援。

3. **使用愉快並充分了解細節**

    人們享受使用功能性與外觀兼具的產品。您可以精心打造您已仔細考量每個互動和視覺效果細節的體驗，以確保您的增益集廣受歡迎。完成工作的必要步驟並需清楚而相關。在理想的情況下，任何動作最好不要超過按一下或兩下滑鼠。盡量不要讓使用者離開完成動作的情境。使用者應該能夠輕易地進出您的增益集，並返回她先前所做的動作。增益集不是要花太多時間的目的地，它可增強我們的核心功能。如果表現得當，增益集可協助我們達成讓使用者更具生產力的目標。

4. **明智發揮品牌價值**

    我們尊重很棒的品牌，而且我們知道一定要為使用者提供您獨特的經驗。但是，我們覺得確保增益集成功的最佳方式，就是建立可巧妙融入您的品牌的直覺體驗，而不是持續顯示突兀的品牌元素，這只會讓使用者分心，而無法在不受妨礙的情況下瀏覽您的系統。運用您的品牌色彩、圖示和代言人，便可以用有意義的方式融入品牌 – 但前提是上述這些項目不會與偏好的平台模式或可及性需求相衝突。盡量將焦點放在內容和工作完成度，而不是品牌關注度。

## <a name="design-patterns"></a>設計模式

> **附註：**雖然上述原則適用於所有的端點/平台，但下列模式和範例則專用於 iOS 平台上的行動裝置增益集。

為了協助您建立一個設計完善的增益集，我們有一些[範本](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/tree/master/Helpful%20Templates/Outlook%20Mobile)，其中包含在 Outlook Mobile 環境中運作的 iOS 行動模式。運用這些特定模式，有助於確保增益集為 iOS 平台和 Outlook Mobile 的原生增益集。這些模式也會詳述如下。雖不詳盡，但這是當我們發現協力廠商想納入其增益集中的其他範例時將持續建置程式庫的開頭。  

### <a name="overview"></a>Overview

典型增益集是由下列元件組成。

![iOS 上工作窗格的基本 UX 模式圖表](../../images/outlook-mobile-design-overview.png)

### <a name="loading"></a>載入

當使用者點選您的增益集時，UX 應盡快顯示。如有任何延遲，請使用進度列或活動指標。進度列應使用於可判定時間量的情況，而活動指標則應使用於無法判定時間量的情況。

![iOS 上的進度列和活動指標範例](../../images/outlook-mobile-design-loading.png)

### <a name="sign-insign-up"></a>登入/註冊

讓您的登入 (和註冊) 流程直接且易於使用。

![iOS 上的登入和註冊頁面範例](../../images/outlook-mobile-design-signin.png)

### <a name="brand-bar"></a>品牌列

增益集的第一個畫面應該包含您的品牌元素。品牌列是為了辨識而設計的，也有助於設定使用者的內容。因為導覽列包含您的公司/品牌名稱，所以後續頁面上不必重複顯示品牌列。

![iOS 上的品牌列範例](../../images/outlook-mobile-design-branding.png)

### <a name="margins"></a>邊界

行動裝置邊界應該設定為每一邊 15px (螢幕的 8%)，以配合 Outlook iOS。

![iOS 上的邊界範例](../../images/outlook-mobile-design-margins.png)

### <a name="typography"></a>印刷樣式

印刷樣式很適合 Outlook iOS，而且樣式簡單、易於掃讀。

![iOS 的印刷樣式範例](../../images/outlook-mobile-design-typography.png)

### <a name="color-palette"></a>調色盤

Outlook iOS 中的色彩用法細緻。為求一致，我們要求色彩用法符合當地化動作和錯誤狀態，只有品牌列使用獨特的色彩。

![iOS 的調色盤](../../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a>儲存格

由於導覽列無法用來標記頁面，所以使用標題來標記頁面。

![iOS 的儲存格類型](../../images/outlook-mobile-design-cell-types.png)
* * *
![iOS 的儲存格可行事項](../../images/outlook-mobile-design-cell-dos.png)
* * *
![iOS 的儲存格不可行事項](../../images/outlook-mobile-design-cell-donts.png)
* * *
![iOS 的儲存格和輸入](../../images/outlook-mobile-design-cell-input.png)

### <a name="actions"></a>動作

即使您的應用程式可處理許多動作，請思考您希望增益集執行的最重要動作，並專注於這些動作。

![iOS 中的動作和儲存格](../../images/outlook-mobile-design-action-cells.png)
* * *
![iOS 的可行動作](../../images/outlook-mobile-design-action-dos.png)

### <a name="buttons"></a>按鈕

有下列其他 UX 元素時使用的按鈕 (相對於動作，其中動作是畫面上的最後一個元素)。

![iOS 的按鈕範例](../../images/outlook-mobile-design-buttons.png)

### <a name="tabs"></a>索引標籤

索引標籤對內容組織方式有所助益。

![iOS 版的索引標籤範例](../../images/outlook-mobile-design-tabs.png)

### <a name="icons"></a>圖示

圖示應儘可能遵循目前的 Outlook iOS 設計。使用我們的標準大小和色彩。

![iOS 的圖示範例](../../images/outlook-mobile-design-icons.png)

## <a name="end-to-end-examples"></a>端對端範例

為了推出我們的 v1 Outlook Mobile 增益集，我們與建置增益集的合作夥伴密切合作。我們的設計人員將每個增益集的端對端流程放在一起，並利用我們的指導方針和模式，以展現其增益集在 Outlook Mobile 上的潛力。

> **重要注意事項︰**這些範例的用意是要強調著手進行理增益集互動與視覺效果設計的理想方式，而可能不符合正式版增益集的確切功能集。 

### <a name="giphy"></a>GIPHY

![GIPHY 增益集的端對端設計](../../images/outlook-mobile-design-giphy.png)

### <a name="nimble"></a>Nimble

![Nimble 增益集的端對端設計](../../images/outlook-mobile-design-nimble.png)

### <a name="trello"></a>Trello

![Trello 增益集第 1 部份的端對端設計](../../images/outlook-mobile-design-trello-1.png)
* * *
![Trello 增益集第 2 部份的端對端設計](../../images/outlook-mobile-design-trello-2.png)
* * *
![Trello 增益集第 3 部份的端對端設計](../../images/outlook-mobile-design-trello-3.png)

### <a name="dynamics-crm"></a>Dynamics CRM

![Dynamics CRM 增益集的端對端設計](../../images/outlook-mobile-design-crm.png)
