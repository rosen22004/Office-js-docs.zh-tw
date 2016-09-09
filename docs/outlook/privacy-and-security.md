
# Outlook 增益集的隱私權、權限和安全性
使用者、開發人員和系統管理員可以使用 Outlook 增益集安全性模型的階梯式權限層級來控制隱私性和效能。



本文說明 Outlook 增益集可以要求的可能權限，並會從下列方面檢查安全性模型︰

- Office 市集 - 增益集完整性
    
- 使用者 - 隱私權與效能考量。
    
- 開發人員 - 權限選項及資源使用量限制。
    
- 系統管理員 - 設定效能臨界值的權限。
    

## 權限模型


因為客戶對於增益集安全性的認知可能會影響增益集的採用，因此 Outlook 增益集安全性依賴分層的權限模型。Outlook 增益集會透露其所需的權限層級，可識別增益集可以在客戶的信箱資料上執行的可能存取及動作。 

資訊清單結構描述版本 1.1 包含四個層級的權限。 


**表 1.增益集權限等級**


|**權限等級**|**Outlook 增益集資訊清單中的值**|
|:-----|:-----|
|限制|受限|
|讀取項目|ReadItem|
|讀寫項目|ReadWriteItem|
|讀寫信箱|ReadWriteMailbox|
權限的四個層級為累計：**讀寫信箱**權限包含**讀寫項目**的權限，**讀取項目**和**受限**，**讀寫項目**包含**讀取項目**和**受限**，以及**讀取項目**權限包含**受限**。圖 1 顯示四個等級的權限，並說明依每一個層級提供給使用者、開發人員和系統管理員的能力。如需有關這些權限的詳細資訊，請參閱[使用者：隱私性和效能考量](#使用者：隱私性和效能考量)、[開發人員︰權限選項及資源使用量限制](#開發人員︰權限選項及資源使用量限制)，以及[了解 Outlook 增益集的權限](../outlook/understanding-outlook-add-in-permissions.md)。 


**圖 1.關於使用者、開發人員和系統管理員的 4 層的權限模型**

![郵件應用程式結構描述 v1.1 的 4 層權限模式](../../images/olowa15wecon_Permissions_4Tier.png)


## Office 市集：增益集完整性


Office 市集裝載可由使用者和系統管理員安裝的增益集。Office 市集強制執行下列措施來保有這些 Outlook 增益集的完整性︰


- 需要增益集的主機伺服器一律使用安全通訊端層 (SSL) 通訊。
    
- 需要開發人員提供的識別證明、契約性協議，以及相容的隱私權原則以送出增益集。 
    
- 在唯讀模式中的保存檔增益集。
    
- 支援可用增益集的使用者檢閱系統，以提升自我監督的社群。
    

## 使用者：隱私權與效能考量。


安全性模型會以下列方式解決安全性、隱私權和使用者的效能問題︰


- 受 Outlook 的資訊版權管理 (IRM) 保護的使用者郵件無法與 Outlook 增益集進行互動。
    
- 從 Office 市集安裝增益集之前，使用者可看見增益集可以在其資料上進行的存取並執行，且必須明確地確認以繼續。Outlook 增益集不會自動推入至不含由使用者或系統管理員手動驗證的用戶端電腦。
    
- 授與**受限**使用權限可以讓 Outlook 增益集僅在目前項目上具有有限的存取。授與**讀取項目**使用權限可以讓 Outlook 增益集存取僅在目前項目上的個人可識別資訊 (例如寄件者和收件者名稱及電子郵件地址)。
    
- 使用者僅可為自己安裝 Outlook 增益集。會影響組織的 Outlook 增益集是由系統管理員進行安裝。
    
- 使用者可以安裝啟用即時線上案例的 Outlook 增益集，這些案例令使用者讚嘆並且將安全性風險降到最低。
    
- 已安裝的 Outlook 增益集的資訊清單檔案會在使用者的電子郵件帳戶中受到保護。
    
- 與裝載 Office 增益集的伺服器通訊的資料一律會根據安全通訊端層 (SSL) 通訊協定加密。
    
- 僅 Outlook 豐富型用戶端適用︰Outlook 豐富型用戶端會監視已安裝的 Outlook 增益集的效能、執行控管控制項，並在下列區域停用這些超出限制的 Outlook 增益集︰
    
      - Response time to activate
    
  - 啟動或重新啟動的失敗次數
    
  - 記憶體使用量
    
  - CPU 使用量
    

    控管會阻擋拒絕服務攻擊，並維持合理的增益集效能。 「商務列」會提供使用者 Outlook 增益集的相關警示：根據此控管控制項，系統已停用 Outlook 豐富型用戶端。
    
- 任何時候，使用者可以驗證已安裝 Outlook 增益集所要求的權限，並停用或接著在 Exchange 系統管理中心啟用任何 Outlook 增益集。
    

## 開發人員：權限選項及資源使用量限制。


安全性模型提供開發人員權限的細微層級以從中選擇，以及嚴格效能方針以進行觀察。


### 分層的權限會增加透明度

開發人員應該依照分層的權限模型來提供透明度，並減輕使用者對於增益集可以對其資料及信箱進行之動作的考量，間接升級增益集採用︰


- 開發人員會根據 Outlook 增益集應該啟動的方式，以及其讀取或寫入的某些項目屬性或建立及傳送項目的需要，來要求適當的 Outlook 增益集權限層級。
    
- 開發人員使用 Outlook 增益集資訊清單中的 [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) 元素來要求使用權限，方法為視需要指派 **Restricted**、**ReadItem**、**ReadWriteItem** 或 **ReadWriteMailbox** 的值。 
    
     >**附註**  請注意，自資訊清單結構描述 v1.1 開始，可以使用 **ReadWriteItem** 權限。

    下列範例要求**讀取項目**權限。
    


```XML
  <Permissions>ReadItem</Permissions>
```

- 如果 Outlook 增益集在特定類型的 Outlook 項目 (約會或郵件) 上，或要在項目的主旨或本文中呈現的特定擷取實體 (電話號碼、地址、URL) 上啟動，則開發人員可以要求**受限**權限。例如，如果在目前郵件的主旨或本文中找到一個或多個以下三個實體 - 電話號碼、郵寄地址或 URL，則下列規則會啟動 Outlook 增益集。
    
```XML
  <Permissions>Restricted</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
    </Rule>
</Rule>
```

- 如果 Outlook 增益集除了預設擷取的實體之外還需要閱讀目前的項目，或藉由目前項目上的增益集撰寫自訂屬性集，但不需要讀取或寫入至其他項目，或是建立或傳送使用者信箱中的郵件，則開發人員應該要求**讀取項目**權限。例如，如果 Outlook 增益集需要在項目的主旨或本文中尋找如會議建議、工作建議、電子郵件地址或連絡人名稱的實體，或使用規則運算式來啟動，則開發人員應該要求**讀取項目**的權限。
    
- 如果 Outlook 增益集需要寫入組成項目的屬性，如收件者名稱、電子郵件地址、本文和主旨，或是需要新增或移除項目附件，則開發人員應該要求**讀寫項目**的權限。
    
- 只有當 Outlook 增益集需要使用 **mailbox.makeEWSRequestAsync** 方法來執行一或多個下列的動作時，開發人員才會要求[讀寫信箱](../../reference/outlook/Office.context.mailbox.md)權限︰
    
      - Read or write to properties of items in the mailbox.
    
  - 建立、讀取、寫入，或傳送信箱中的項目。
    
  - 建立、讀取或寫入信箱中的資料夾。
    

### 資源使用狀況調整

開發人員應該注意啟動的資源使用量限制，在其開發工作流程中合併效能調整，以減少執行效能不佳的增益集拒絕主應用程式服務的機會。開發人員在設計[適用於 Outlook 增益集的 JavaScript API 和啟動的限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)中所述的啟動規則時，應該遵循指導方針。如果 Outlook 增益集要在 Outlook 豐富型用戶端上執行，開發人員應該確認增益集在資源使用量限制內執行。


### 提升使用者安全性的其他措施

開發人員也應該注意並計劃下列項目︰


- 開發人員無法使用增益集中的 ActiveX 控制項，因為其不受支援。
    
- 開發人員提交增益集至 Office 市集時，應該執行下列動作︰
    
      - Produce an Extended Validation (EV) SSL certificate as a proof of identity.
    
  - 在支援 SSL 的網頁伺服器上主控他們要提交的增益集。
    
  - 產生相容的隱私權原則。
    
  - 準備在送出增益集時簽署契約性協議。
    

## 系統管理員︰權限


安全性模型提供下列權限及責任給系統管理員︰


- 可以防止使用者安裝任何 Outlook 增益集，包括 Office 市集上的增益集。
    
- 可以在 Exchange 系統管理中心停用或啟用任何 Outlook 增益集。
    
- 僅適用 Outlook for Windows:可以依 GPO 登錄設定覆寫效能臨界值設定。
    


## 其他資源



- [Outlook 增益集](../outlook/outlook-add-ins.md)
    
- [Office 增益集的隱私權和安全性](../../docs/develop/privacy-and-security.md)
    
- [Outlook 增益集 API](../outlook/apis.md)
    
- [要求用於內容和工作窗格增益集的 API 權限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
    
- [適用於 Outlook 增益集的 JavaScript API 和啟動的限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
