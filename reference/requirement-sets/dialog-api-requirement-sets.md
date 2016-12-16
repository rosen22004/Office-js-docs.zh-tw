
# <a name="dialog-api-requirement-sets"></a>對話方塊 API 需求集合

需求集合是 API 成員的具名群組。Office 增益集使用資訊清單中所指定的需求集合，或使用執行階段檢查，以判定 Office 主應用程式是否支援增益集所需的的 API。如需詳細資訊，請參閱[指定 Office 主應用程式及 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

Office 增益集執行多個版本的 Office。下表列出對話方塊 API 需求集合、支援需求集合的 Office 主應用程式，以及 Office 應用程式的組建或版本號碼。

|  需求集合  |  Office 2013 for Windows | Office 2016 for Windows*   |  Office 2016 for iPad  |  Mac 版 Office 2016  | Office Online  |  Office Online 伺服器  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | 組建 15.0.4855.1000 或更新版本 | 版本 1602 (組建 6741.0000) 或更新版本 | 1.22 或更新版本 | 15.20 或更新版本| 我們正在完成它。 | 版本 1608 (組建 7601.6800) 或更新版本|

>**附註：**透過 MSI 安裝的 Office 2016 組建編號是 16.0.4266.1001。若要使用對話方塊 API，請執行 Office 更新以取得最新版本。 

若要瞭解關於版本、組建編號及 Office Online Server 的詳細資訊，請參閱︰

- [Office 365 用戶端的更新通道版本和組建編號](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [我使用的是哪個版本的 Office？](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [您可以在其中找到 Office 365 用戶端應用程式的版本和組建編號](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server 概觀](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 需求集合
如需通用 API 需求集合的詳細資訊，請參閱 [Office 通用 API 需求集合](office-add-in-requirement-sets.md)。

## <a name="dialog-api-11"></a>對話方塊 API 1.1 
對話方塊 API 1.1 是 API 的第一個版本。如需 API 的詳細資訊，請參閱[對話方塊 API](../shared/officeui.md) 參考主題。

## <a name="additional-resources"></a>其他資源

- [指定 Office 主應用程式和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office 增益集的 XML 資訊清單](../../docs/overview/add-in-manifests.md)

