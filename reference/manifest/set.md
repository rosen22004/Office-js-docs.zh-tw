
# <a name="set-element"></a>Set 項目
指定適用於 Office 的 JavaScript API 中，為了啟用您的 Office 增益集所需的需求集合。

 **增益集類型︰**內容、工作窗格、郵件


## <a name="syntax:"></a>語法：


```XML
<Set Name="string " MinVersion="n .n ">
```


## <a name="contained-in:"></a>內含於：

[Sets](../../reference/manifest/sets.md)


## <a name="attributes"></a>屬性



|**屬性**|**類型**|**必要**|**描述**|
|:-----|:-----|:-----|:-----|
|名稱|字串|必要|[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) 的名稱。|
|MinVersion|String|選用|指定增益集所需的 API 集合最小版本。如果已在父 **Sets** 項目中將其指定，就會覆寫 [DefaultMinVersion](../../reference/manifest/sets.md) 的值。|

## <a name="remarks"></a>備註

如需有關需求集合的詳細資訊，請參閱[指定 Office 主機和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md#specify-office-hosts-and-api-requirements)。

如需有關 **Set** 項目的 **MinVersion** 屬性和 **Sets** 項目的 **DefaultMinVersion** 屬性的詳細資訊，請參閱[指定 Office 主機和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。


 >**重要**  對郵件增益集，只有一個 `"Mailbox"` 需求集合可用。這個需求集合包含 Outlook 的郵件增益集中所支援的整個 API 子集，而且您必須在郵件增益集的資訊清單中指定 `"Mailbox"` 需求集合 (對內容和工作窗格增益集不是選用)。此外，您無法宣告支援郵件增益集中的特定方法。

