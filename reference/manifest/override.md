
# Override 項目
提供指定其他的地區設定的設定值的方法。

 **增益集類型︰**內容、工作窗格、郵件


## 語法：


```XML
<Override Locale="string " Value="string " />
```


## 內含於：


||
|:-----|
|[CitationText](../../reference/manifest/citationtext.md)|
|[說明](../../reference/manifest/description.md)|
|[DictionaryName](../../reference/manifest/dictionaryname.md)|
|[DictionaryHomePage](../../reference/manifest/dictionaryhomepage.md)|
|[DisplayName](../../reference/manifest/displayname.md)|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|
|[IconUrl](../../reference/manifest/iconurl.md)|
|[QueryUri](../../reference/manifest/queryuri.md)|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|
|[SupportUrl](../../reference/manifest/supporturl.md)|

## 屬性



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|地區設定|String|必要|以 BCP 47 語言標記格式，指定這個覆寫的地區設定文化特性名稱，例如 `"en-US"`。|
|值|string|必要|為指定之文化特性所表示的設定指定值。|

## 其他資源



- [Office 增益集的當地語系化](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
