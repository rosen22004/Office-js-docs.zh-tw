
# <a name="context.mailbox-property"></a>Context.mailbox 屬性
取得 **mailbox** 物件，其可供存取專用於 Outlook 增益集的 API 成員。

|||
|:-----|:-----|
|**主應用程式︰**|Outlook|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|信箱|
|**上次變更於**|1.0|

```js
var outlookOm = Office.context.mailbox;
```


## <a name="return-value"></a>傳回值


  [mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx) 物件。


## <a name="example"></a>範例

下列這一行程式碼會存取 JavaScript API for Office 的 [item](http://msdn.microsoft.com/library/ad288df1-3ca2-474c-bea4-c51f46e6fc43%28Office.15%29.aspx) 物件。


```js
// Access the Item object.
var item = Office.context.mailbox.item;

```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|信箱|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|Outlook|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.0|已導入|
