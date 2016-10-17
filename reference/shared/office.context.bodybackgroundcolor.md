
# <a name="officetheme.bodybackgroundcolor-property"></a>officeTheme.bodyBackgroundColor 屬性
取得 Office 佈景主題內容背景色彩。

 **重要事項：**此 API 目前只能在 Windows 桌面上，用於 [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) 的 Excel、Outlook、PowerPoint 和 Word 中。


|||
|:-----|:-----|
|**主應用程式︰**|Excel、Outlook、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|不在集合中|
|**已新增於**|1.3|



```
var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
```


## <a name="return-value"></a>傳回值

十六進位彩色三元組。


## <a name="remarks"></a>備註

傳回的色彩對應於使用者透過 **[檔案]**  >  **[Office 帳戶]**  >  **[Office 佈景主題]** UI 所選取的 Office 佈景主題值 (套用於所有 Office 主應用程式)。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支援的主應用程式 (依平台排序)**


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||||
|**Outlook**|Y||||
|**PowerPoint**|Y||||
|**Word**|Y||||

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.3|已導入|
