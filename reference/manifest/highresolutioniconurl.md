
# HighResolutionIconUrl 項目
指定在高 DPI 的螢幕上，插入 UX 和 Office 市集中，用來表示您 Office 增益集的影像 URL。

 **增益集類型︰**內容、工作窗格、郵件


## 語法：


```XML
<HighResolutionIconUrl DefaultValue="string " />
```


## 可以包含︰

[Override](../../reference/manifest/override.md)


## 屬性



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string (URL)|必要|指定此設定的預設值，這個值會在 [DefaultLocale](../../reference/manifest/defaultlocale.md) 項目中指定的地區設定內顯示。|

## 備註

對郵件增益集，圖示會顯示在 [檔案]**File** > **[管理增益集]** UI。 對內容或工作窗格增益集，圖示會顯示在 [插入]****  > ** [增益集]** UI。

影像必須是下列檔案格式之一，建議使用解析度 64 x 64 像素：GIF、JPG、PNG、EXIF、BMP 或 TIFF。 如需詳細資訊，請參閱[建立有效的 Office 市集應用程式和增益集](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)中的_＜為您的應用程式建立一致的視覺識別＞_章節。

