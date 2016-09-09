
# Host 項目
指定您 Office 增益集所支援的 Office 主應用程式類型。

 **增益集類型︰**內容、工作窗格、郵件


## 語法：


```XML
<Host Name= ["Document" | "Database" | "Mailbox" | "Presentation" | "Project" | "Workbook"] />
```


## 屬性



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|名稱|string|必要|Office 主應用程式的類型名稱。|

## 備註

您可以在 **Host** 項目的 **Name** 屬性中指定下列值。每個值會對應到您的增益集所支援的一或多個 Office 主應用程式集。



|**名稱**|**Office 主應用程式**|
|:-----|:-----|
| `"Document"`|Word、Word Online、iPad 版 Word|
| `"Database"`|Access Web App|
| `"Mailbox"`|Outlook、Outlook Web App、裝置用 OWA|
| `"Notebook"`|OneNote Online|
| `"Presentation"`|PowerPoint、PowerPoint Online、iPad 版 PowerPoint|
| `"Project"`|Project|
| `"Workbook"`|Excel、Excel Online、iPad 版 Excel|

## 備註

如需有關如何指定主機支援的詳細資訊，請參閱[指定 Office 主機和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

