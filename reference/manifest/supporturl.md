
# <a name="supporturl-element"></a>SupportUrl 項目

指定提供您增益集的支援資訊的網頁 URL。

## <a name="example"></a>範例

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>

```


## <a name="attributes"></a>屬性

|**屬性**|**類型**|**必要**|**描述**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必要|指定此設定的預設值，這個值會在 [DefaultLocale](../../reference/manifest/defaultlocale.md) 項目中指定的地區設定內顯示。|

## <a name="child-elements"></a>子元素

|  元素 | 必要 | 描述  |
|:-----|:-----|:-----|
|  [Override](../../reference/manifest/override.md)   | 否 | 指定其他地區設定 URL 的設定 |

## <a name="parent-element"></a>父元素
[OfficeApp](../../reference/manifest/officeapp.md)

