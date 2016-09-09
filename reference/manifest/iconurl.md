
# IconUrl 項目
指定在插入 UX 和 Office 市集中，用來表示您 Office 增益集的影像 URL。

 **增益集類型︰**內容、工作窗格、郵件


## 語法：


```XML
<IconUrl DefaultValue="string " />
```


## 可以包含︰

[Override](../../reference/manifest/override.md)


## 屬性



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|String|必要|指定此設定的預設值，這個值會在 [DefaultLocale](../../reference/manifest/defaultlocale.md) 項目中指定的地區設定內顯示。|

## 備註

對郵件增益集，圖示會顯示在 [檔案]**** >  [管理增益集]**** UI (Outlook) 或 [設定]****  >  [管理增益集]**** UI (Outlook Web App)。 對內容或工作窗格增益集，圖示會顯示在 [插入]****  > ** [增益集]** UI。 對於所有的增益集類型，如果您將增益集發佈到 Office 市集，則圖示也會用於 Office 市集網站上。

影像必須是下列檔案格式之一︰GIF、JPG、PNG、EXIF、BMP 或 TIFF。對內容和工作窗格增益集，指定的影像必須是 32 x 32 像素。對郵件增益集，影像必須是 64 x 64 像素。您也應該指定圖示，以便在高 DPI 的螢幕上執行 Office 主應用程式時，使用 [HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md) 項目。如需詳細資訊，請參閱_建立有效的 Office 市集應用程式和增益集_中的[＜為您的應用程式建立一致的視覺識別＞](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)章節。

