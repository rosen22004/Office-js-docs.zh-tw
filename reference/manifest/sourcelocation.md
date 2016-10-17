
# <a name="sourcelocation-element"></a>SourceLocation 項目
將您 Office 增益集的來源檔位置，指定為 URL，長度介於 1 到 2018 個字元。來源位置必須是 HTTPS 位址，不得是檔案路徑。

 **增益集類型：**內容、工作窗格、郵件


## <a name="syntax:"></a>語法：


```XML
<SourceLocation DefaultValue="string " />
```


## <a name="contained-in:"></a>內含於：

[DefaultSettings](../../reference/manifest/defaultsettings.md) (內容和工作窗格增益集)

[FormSettings](../../reference/manifest/formsettings.md) (郵件增益集)


## <a name="can-contain:"></a>可以包含︰

[Override](../../reference/manifest/override.md)


## <a name="attributes"></a>屬性



|**屬性**|**類型**|**必要**|**描述**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必要|指定在 [DefaultLocale](../../reference/manifest/defaultlocale.md) 項目中指定的地區設定的此設定預設值。|
