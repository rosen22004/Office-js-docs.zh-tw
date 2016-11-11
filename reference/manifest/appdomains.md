
# <a name="appdomains-element"></a>AppDomains 元素
除了 [SourceLocation](../../reference/manifest/sourcelocation.md) 元素中指定的網域之外，列出您的 Office 增益集會用來載入頁面的所有網域。針對每個其他網域，指定 [AppDomain](../../reference/manifest/appdomain.md) 元素。

 **增益集類型：**內容、工作窗格、郵件


## <a name="syntax"></a>語法：


```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```


## <a name="contained-in"></a>內含於：

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain"></a>可以包含︰

[AppDomain](../../reference/manifest/appdomain.md)


## <a name="remarks"></a>備註

依預設，您的增益集可以載入相同網域中的任何頁面，如 [SourceLocation](../../reference/manifest/sourcelocation.md) 元素中所指定的。若要載入不同網域中的頁面做為增益集，請使用 **AppDomains** 和 **AppDomain** 元素來指定網域此元素不能為空。 

如需詳細資訊，請參閱 [Office 增益集 XML 資訊清單](../../docs/overview/add-in-manifests.md)。

