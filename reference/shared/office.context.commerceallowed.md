
# <a name="context.commerceallowed-property"></a>Context.commerceAllowed 屬性
取得關於增益集是否在允許連結到外部付款系統的平台上執行的資訊。

|||
|:-----|:-----|
|**主應用程式︰**|Excel、Word|
|**上次變更於**|1.1|

```
var allowCommerce = Office.context.commerceAllowed;
```


## <a name="return-value"></a>傳回值

如果開發人員可以在該平台上的增益集中顯示銷售或升級 UI，則傳回 **True**；否則傳回 **False**。


## <a name="remarks"></a>備註

iOS App Store 不支援含有提供其他付款系統連結之增益集的應用程式。不過，在 Windows 桌面上執行的 Office 增益集，或是在瀏覽器中用於 Office Online 的 Office 增益集，可允許這類連結。如果您想要讓增益集的 UI 在 iOS 以外的平台上，提供外部付款系統的連結，您可以使用 **commerceAllowed** 屬性來控制何時顯示該連結。


## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for iPad**|
|:-----|:-----|
|**Excel**|Y|
|**PowerPoint**||
|**Word**|Y|

|||
|:-----|:-----|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.1|已導入。|
