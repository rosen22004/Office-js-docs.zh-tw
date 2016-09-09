

# Settings.remove 方法
移除指定的設定。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、PowerPoint、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**上次變更於**|1.1|

```js
Office.context.document.settings.remove(name);
```


## 參數



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;類型：**字串**

&nbsp;&nbsp;&nbsp;&nbsp;要移除之設定的區分大小寫名稱。

    



## 備註

 **null** 是有效的設定值。因此，指派 **null** 至設定不會將該設定從設定屬性包移除。


 >**重要**：請注意，**Settings.remove** 方法僅影響設定屬性包的記憶體內部複本。若要在文件中保存移除指定設定，您必須在呼叫 **Settings.remove** 方法後及關閉增益集前這段時間內，呼叫 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法。


## 範例




```js
function removeMySetting() {
    Office.context.document.settings.remove('mySetting');
}
```




## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**可用於需求集合**|Settings|
|**最低權限等級**|[限制](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄




|**版本**|**變更**|
|:-----|:-----|
|1.1|新增對 PowerPoint Online 的支援。|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|新增支援在 Access 內容增益集中自訂設定。|
|1.0|已導入|
