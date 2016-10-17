
# <a name="custom-pane-outlook-add-ins"></a>自訂窗格 Outlook 增益集

自訂窗格是增益集的擴充點，當目前選取的項目上符合特定條件時會啟動。其在 **VersionOverrides** 元素的增益集資訊清單中進行定義，連同任何由增益集所實作的增益集命令。如需詳細資訊，請參閱[定義 Outlook 增益集資訊清單中的增益集命令](../outlook/manifests/define-add-in-commands.md)。自訂窗格只能出現在讀取的郵件或約會出席者檢視中。它會在增益集列中顯示一個項目。當使用者按一下項目時，自訂窗格會以水平方向顯示在項目本文的上方。外觀和行為與不會實作增益集命令的讀取模式增益集相同。

**具有讀取模式中自訂窗格的增益集**

![顯示郵件閱讀表單中的自訂窗格。](../../images/c585ab0a-6c33-42d0-a20f-5deb8b54f480.png)

下列範例會定義屬於郵件或具有附件或包含位址的項目的自訂窗格。 



```
<ExtensionPoint xsi:type="CustomPane">
  <RequestedHeight>100< /RequestedHeight> 
  <SourceLocation resid="residReadTaskpaneUrl"/>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasAttachment"/>
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
  </Rule>
</ExtensionPoint>
```



-  在桌上型電腦上執行時，**RequestedHeight** 會指定此郵件增益集的所需高度 (以像素為單位)。否則會忽略。它可以是一個介於 32 到 450 的值。如果沒有設定，預設值為 350 像素：選用。
    
-  **SourceLocation** 會指定提供自訂窗格的 UI 的 HTML 頁面。**resid** 屬性會設定為 **資源**元素中 **URL** 元素的 **id** 屬性的值。必要。
    
-  
  **Rule** 指定增益集啟動時指定的規則或規則集合。其與 [Outlook 增益集資訊清單](../outlook/manifests/manifests.md)中所定義的相同，除了 [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) 規則具有以下變更：**ItemType** 是「Message」或「AppointmentAttendee」，且沒有 **FormType** 屬性。如需詳細資訊，請參閱 [Outlook 增益集的啟用規則](../outlook/manifests/activation-rules.md)。
    

## <a name="additional-resources"></a>其他資源



- [開始使用 Office 365 的 Outlook 增益集](https://dev.outlook.com/MailAppsGettingStarted)
    
- [Outlook 增益集的啟用規則](../outlook/manifests/activation-rules.md)
    
- [Outlook 增益集資訊清單](../outlook/manifests/manifests.md)
    
- [在 Outlook 增益集資訊清單中定義增益集命令](../outlook/manifests/define-add-in-commands.md)
    
