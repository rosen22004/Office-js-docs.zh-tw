
# <a name="officeapp-element"></a>OfficeApp 項目
Office 增益集資訊清單中的根項目。

 **增益集類型︰**內容、工作窗格、郵件


## <a name="syntax:"></a>語法：


```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```


## <a name="contained-in:"></a>內含於：

 _無_


## <a name="must-contain:"></a>必須包含︰



|** 元素**|**內容**|**郵件**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Id](../../reference/manifest/id.md)|x|x|x|
|[Version](../../reference/manifest/version.md)|x|x|x|
|[ProviderName](../../reference/manifest/providername.md)|x|x|x|
|[DefaultLocale](../../reference/manifest/defaultlocale.md)|x|x|x|
|[DefaultSettings](../../reference/manifest/defaultsettings.md)|x|x|x|
|[DisplayName](../../reference/manifest/displayname.md)|x|x|x|
|[Description](../../reference/manifest/description.md)|x|x|x|
|[FormSettings](../../reference/manifest/formsettings.md)||x||
|[Permissions](../../reference/manifest/permissions.md)|x||x|
|[Rule](../../reference/manifest/rule.md)||x||

## <a name="can-contain:"></a>可以包含︰



|** 元素**|**內容**|**郵件**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](../../reference/manifest/alternateid.md)|x|x|x|
|[IconUrl](../../reference/manifest/iconurl.md)|x|x|x|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|x|x|x|
|[SupportUrl](../../reference/manifest/supporturl.md)|x|x|x|
|[AppDomains](../../reference/manifest/appdomains.md)|x|x|x|
|[Hosts](../../reference/manifest/hosts.md)|x|x|x|
|[Requirements](../../reference/manifest/requirements.md)|x|x|x|
|[AllowSnapshot](../../reference/manifest/allowsnapshot.md)|x|||
|[Permissions](../../reference/manifest/permissions.md)||x||
|[DisableEntityHighlighting](../../reference/manifest/disableentityhighlighting.md)||x||
|[Dictionary](../../reference/manifest/dictionary.md)|||x|
|[VersionOverrides](../../reference/manifest/versionoverrides.md)|X|X|X|

## <a name="attributes"></a>屬性


|||
|:-----|:-----|
|xmlns|定義 Office 增益集的資訊清單命名空間和結構描述版本。這個屬性應該永遠設定為 `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|定義 XMLSchema 執行個體。這個屬性應該永遠設定為 `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|定義 Office 增益集的種類。這個屬性應該設為下列其中一項︰`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`|
