
# 部署和安裝 Outlook 增益集以進行測試


做為開發 Outlook 增益集程序的一部分，您可能會發現自己反覆部署並安裝增益集來進行測試，其中包含下列步驟︰


1. 建立說明增益集的資訊清單檔。
    
2. 部署增益集 UI 檔案至 web 伺服器。
    
3. 在您的信箱中安裝增益集。
    
4. 測試增益集，對 UI 或資訊清單檔進行適當變更，並重複步驟 2 和 3 以測試所做的變更。
    

## 建立增益集的資訊清單檔

每個增益集是由 XML 資訊清單所描述，其為提供增益集相關資訊、為使用者提供關於增益集的描述性資訊，以及識別增益集 UI HTML 檔案位置的文件。您可以在本機資料夾或伺服器中儲存資訊清單，只要您用來測試的信箱的 Exchange Server 可以存取位置。我們會假設您在本機資料夾中儲存資訊清單。如需有關如何建立資訊清單檔的資訊，請參閱 [Outlook 增益集資訊清單](../outlook/manifests/manifests.md)。 


## 部署增益集至 web 伺服器

您可以使用 HTML 和 JavaScript 來建立增益集 UI。產生的原始程式檔會儲存在可由裝載增益集的 Exchange Server 存取的 web 伺服器上。原始程式檔會由 **DesktopSettings** 元素、[TabletSettings](http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx) 元素和 (或) 資訊清單檔中指定的 [PhoneSettings](http://msdn.microsoft.com/en-us/library/5c89cc7c-7ae0-49c9-fdd5-4c52118228f6%28Office.15%29.aspx) 元素中的 [SourceLocation](http://msdn.microsoft.com/en-us/library/13e4eae3-8e8c-fd55-a1c2-3297b485f327%28Office.15%29.aspx) 子元素所識別。

在最初部署增益集的 UI 檔案之後，您可以使用新版本的 HTML 檔案取代儲存在 web 伺服器上的 HTML 檔案，以更新增益集 UI 和行為。


## 安裝增益集


在準備增益集資訊清單檔並部署增益集 UI 至可以存取的 web 伺服器之後，您可以在 Exchange Server 上安裝信箱的增益集，方法為使用 Outlook 豐富型用戶端、Outlook Web App 或裝置用 OWA，或執行遠端 Windows PowerShell cmdlet。


### 在 Outlook 豐富型用戶端中安裝增益集

如果您的信箱位於 Exchange Online、Exchange 2013 或更新版本上，您可以安裝增益集。 在 Outlook for Windows 中，您可以透過 Office Fluent Backstage 檢視安裝增益集。 選擇 [檔案]**** 和 [管理增益集]****。 這可讓您登入 Exchange 系統管理中心。 登入後，繼續下一節的步驟 4 中的安裝程序。

在 Outlook for Mac 中，選擇增益集列右端的 [管理增益集]**** 然後登入 Exchange 系統管理中心。 繼續進行下一節中的步驟 4。


### 使用 Outlook Web App 或 Outlook.com 安裝增益集

若要使用 Outlook Web App (OWA) 安裝 Outlook 增益集，請依照下列步驟執行︰


1. 瀏覽至您組織的 OWA URL 或 Outlook.com 並登入。
    
2. 在右上角選擇齒輪圖示，然後選擇 [管理增益集]****。
    
3. 選取加號 ( **+**) 來新增新的增益集。
    
4. 從下拉式清單中，選取 [從檔案新增]****，假設您已本機資料夾上儲存資訊清單。
    
5. 瀏覽至資訊清單的檔案路徑，然後選取 [安裝]****。
    
6. 選取視窗右上角的使用者名稱，然後選取 [我的郵件]**** 以切換到您的電子郵件來測試增益集。
    

>**附註**  如果您不使用下列其中一項來開發增益集︰ 
- Office 365 開發人員租用戶
- Napa Office 365 Development Tools
- Visual Studio

而且，如果您的 Exchange Server 沒有至少「我的自訂應用程式」角色，則您僅可以從 Office 市集安裝增益集。 為了測試增益集，或藉由指定增益集資訊清單的 URL 或檔案名稱安裝一般增益集，您應該要求您的 Exchange 系統管理員提供必要的權限。

Exchange 系統管理員可以執行下列 PowerShell cmdlet 為單一使用者指派所需的權限。在這個範例中，wendyri 是使用者的電子郵件別名。

```New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"```

如必要，系統管理員可以執行下列 cmdlet，為多個使用者指派類似的必要權限︰

```$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}```

如需「我的自訂應用程式」角色的詳細資訊，請參閱[我的自訂應用程式角色](http://technet.microsoft.com/en-us/library/aa0321b3-2ec0-4694-875b-7a93d3d99089%28EXCHG.150%29.aspx)。 

使用 Office 365、Napa 或 Visual Studio 來開發增益集會為您指派組織系統管理員角色，讓您可藉由檔案或 EAC 中的 URL，或藉由 Powershell cmdlets 安裝增益集。


### 使用遠端 PowerShell 安裝增益集

您在 Exchange server 上建立遠端 Windows PowerShell 工作階段之後，可以使用 **New-App** cmdlet 搭配下列 PowerShell 命令安裝 Outlook 增益集。


```
New-App -URL:"http://<fully-qualified URL">
```

完整的 URL 是您為增益集準備好的增益集資訊清單檔的位置。

您可以使用下列的額外 PowerShell cmdlets 來管理信箱的增益集︰


-  **Get-App** - 列出針對信箱啟用的增益集。
    
-  **Set-App** - 在信箱上啟用或停用增益集。
    
-  **Remove-App** - 從 Exchange Server 移除先前安裝的增益集。
    

## 其他資源



- [Outlook 增益集](../outlook/outlook-add-ins.md)
    
- [疑難排解 Office 增益集的使用者錯誤](../testing/testing-and-troubleshooting.md)
    
