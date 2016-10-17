# <a name="design-guidelines-for-office-add-ins"></a>Office 增益集的設計指導方針

Office 增益集提供可讓使用者在 Office 用戶端中存取的內容相關功能，藉以延伸 Office 經驗。增益集能讓使用者在 Office 中存取協力廠商功能，無需昂貴的內容切換就能完成更多任務。 

 您的增益集 UX 設計必須與 Office 完美整合，才能為使用者提供有效率且自然的互動。建立自訂的 HTML 架構 UI 時，請利用增益集命令 (Office UI 擴充功能) 來提供對您增益集的存取，並使用我們建議的 [UI 元素](ui-elements/ui-elements.md)和[最佳做法](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)。 
 
 
## <a name="core-office-add-in-design-principles"></a>核心 Office 增益集設計原則
不論您用來建立自訂 UI 的基礎架構為何，設計增益集時都請遵循下列原則： 

- **針對 Office 進行明確設計**。增益集的功能以及外觀和風格必須與 Office 經驗產生和諧互補，包括套用 Office 或文件佈景主題。
 
- **讓使用者更有效率**。協助使用者在不干擾其他工作的情況下完成工作。允許 Office 文件和增益集之間順暢互動。 

- **透過組件區塊強化內容**。透過任何附屬組件區塊，強調增益集內容及功能。避免無法提升使用者經驗的多餘 UI 元素，最大化使用空間。  

- **讓使用者能掌控操作**。讓使用者可以掌控其經驗、了解重要的決定，並能輕鬆回復增益集執行的動作。 

- 
  **為所有平台和輸入方法進行設計**。增益集應設計為可在 Office 支援的所有平台上運作，增益集 UX 也應最佳化為可在各種平台和外觀尺寸上運作。支援滑鼠/鍵盤和觸控輸入裝置，並確保自訂 HTML UI 能快速回應以適應不同的外觀尺寸。如需詳細資訊，請參閱[觸控](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#bk_Touch)。 


## <a name="design-language"></a>設計語言
建議您在增益集內採用 Office 設計語言，並使用我們的 [Office UI Fabric](https://dev.office.com/fabric) 來建立自訂的 HTML 架構經驗。如果您的組織已具備設計語言，只要最終產生的成果可為 Office 使用者提供和諧經驗，當然也歡迎您使用。 


## <a name="add-in-building-blocks"></a>增益集建置組塊
您可以使用兩種 UI 元素類型來建立增益集： 

- [增益集命令](ui-elements/ui-elements.md#add-in-commands)可讓您將原生 UX 勾點新增至 Office 應用程式
- [自訂 HTML 架構 UI](ui-elements/ui-elements.md#custom-html-based-ui) 可讓您在 Office 用戶端中運用 HTML 的優點。 

如需如何使用這些建置組塊的詳細資訊，請參閱 [UI 元素](ui-elements/ui-elements.md)。  

## <a name="ux-design-patterns"></a>UX 設計模式

為了協助您建立增益集的第一級使用者經驗，我們會提供說明常見 UX 設計模式的範本。這些範本反映[最佳作法](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)，建立令人讚嘆、世界級的增益集，並包含初次執行體驗、商標元素，以及使用者通知的的模式。他們使用 [Office UI 結構](https://dev.office.com/fabric)元件與樣式，而且包括會自然擴充 Office UI 的元素。

若要存取範本，請參閱 [Office 增益集的 UX 設計模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns) repo。Adobe Illustrator 檔案也是可用的；您可以下載並更新它們以反映出您自己的設計。您也可以從 [Office 增益集的 UX 設計模式程式碼](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo，將程式碼檔案複製到增益集專案，並視需要自訂它們。 

## <a name="recommended-layouts-and-interaction-patterns"></a>建議的配置和互動模式
我們會針對每個增益集類型提供建議的配置以及**端對端**範例，協助您順利完成作業。若要深入了解如何配置增益集，請參閱下列主題：

- [工作窗格容器的版面配置](ui-elements/layout-for-task-pane-add-ins.md)
- [內容增益集的版面配置](ui-elements/layout-for-content-add-ins.md) 
- [郵件增益集的版面配置](ui-elements/layouts-for-outlook-add-ins.md)

另請參閱互動模式，了解增益集的常見案例範例及其對應的互動模式。

## <a name="additional-resources"></a>其他資源

- [Office UI 結構](https://dev.office.com/fabric) 

