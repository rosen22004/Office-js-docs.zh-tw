
# <a name="context.officetheme-property"></a>Context.officeTheme 屬性
可供存取 Office 佈景主題色彩的屬性。

 **重要事項：**此 API 目前只能在 Windows 桌面上，用於 [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) 的 Excel、Outlook、PowerPoint 和 Word 中。


|||
|:-----|:-----|
|**主應用程式︰**|Excel、Outlook、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|不在集合中|
|**已新增於**|1.3|



```js
Office.context.officeTheme
```


## <a name="members"></a>成員


**屬性**

|||
|:-----|:-----|
|名稱|描述|
|[bodyBackgroundColor ](../../reference/shared/office.context.bodybackgroundcolor.md)|取得 Office 佈景主題內容背景色彩。|
|[bodyForegroundColor](../../reference/shared/office.context.bodyforegroundcolor.md)|取得 Office 佈景主題內容前景色彩。|
|[controlBackgroundColor](../../reference/shared/office.context.controlbackgroundcolor.md)|取得 Office 佈景主題控制項的背景色彩。|
|[controlForegroundColor](../../reference/shared/office.context.controlforegroundcolor.md)|取得 Office 佈景主題控制項的前景色彩。|

## <a name="remarks"></a>備註

使用 Office 佈景主題色彩，可讓您針對增益集的色彩配置以及使用者透過 **[檔案]**  >  **[Office 帳戶]**  >  **[Office 佈景主題]** UI 所選取的現行 Office 佈景主題 (套用於所有 Office 主應用程式)，進行協調。Office 佈景主題色彩適用於 Outlook 和工作窗格增益集。


## <a name="example"></a>範例


```js
function applyOfficeTheme(){
    // Get office theme colors.
    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply body background color to a CSS class.
    $('.body').css('background-color', bodyBackgroundColor);
}
```


## <a name="support-details"></a>支援詳細資料



|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格、Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.3|已導入|
