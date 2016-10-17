

# <a name="settings.set-method"></a>Settings.set 方法
設定或建立指定的設定。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|設定|
|**上次變更於**|1.1|

```js
Office.context.document.settings.set(name, value);
```


## <a name="parameters"></a>參數



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**字串**

&nbsp;&nbsp;&nbsp;&nbsp;要設定或建立設定的區分大小寫名稱。

    
_value_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型︰**字串**、**數值**、**布林值**、**null**、**物件**或**陣列**

&nbsp;&nbsp;&nbsp;&nbsp;指定要儲存的值。
    

## <a name="remarks"></a>備註

如果沒有設定，**set** 方法會建立具有指定名稱的新設定，或在設定屬性包的記憶體內部複本中，設定具有指定名稱的現有設定。呼叫 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法後，傳回值會儲存在文件中，成為其資料類型的序列化 JSON 表示法。每個增益集的設定最多可使用 2MB。


 >**重要**：請注意，**Settings.set** 方法只會影響設定屬性包的記憶體內部複本。若要確保下次開啟文件時，增益集可使用新增或變更的設定，您必須在呼叫 **Settings.set** 方法後及關閉增益集前這段時間內，呼叫 **Settings.saveAsync** 方法，將設定保存於文件中。


## <a name="example"></a>範例




```js
function setMySetting() {
    Office.context.document.settings.set('mySetting', 'mySetting value');
}

```




## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**可用於需求集合**|Settings|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增支援在 Access 內容增益集中自訂設定。|
|1.0|已導入|
