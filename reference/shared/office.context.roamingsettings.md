
# Context.roamingSettings 屬性
取得代表自訂設定或儲存至使用者信箱之 Outlook 增益集狀態的物件。

|||
|:-----|:-----|
|**主機︰**|Outlook|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|信箱|
|**上次變更於**|1.0|

```
var appSettings = office.context.roamingSettings;
```


## 傳回值


  [RoamingSettings](http://msdn.microsoft.com/library/cf21bb08-7274-4ad6-ae9e-b2c12f92abc9%28Office.15%29.aspx) 物件。


## 備註

**RoamingSettings** 物件可讓您儲存及存取 Outlook 增益集儲存在使用者信箱中的資料，如此當從用於存取該信箱的任何主機用戶端應用程式上執行增益集時，便可讓其使用該資料。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Mac 版 Outlook**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|信箱|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄



****


|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|
