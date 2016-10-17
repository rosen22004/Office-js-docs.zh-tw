
# <a name="bindings-object"></a>繫結物件
代表增益集在文件中所具有的繫結。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Word|
|**上次變更**於|1.1|

```js
Office.context.document.bindings
```


**屬性**

|||
|:-----|:-----|
|名稱|描述|
|[document](../../reference/shared/bindings.document.md)|取得代表與此組繫結相關聯之文件的 **Document** 物件。|

**方法**

|||
|:-----|:-----|
|名稱|描述|
|[addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)|將繫結加入至文件中的具名項目。|
|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|顯示 UI，讓使用者可指定要繫結的選取範圍。|
|[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|加入繫結至文件中目前選取範圍所指定的類型的繫結物件。|
|[getAllAsync](../../reference/shared/bindings.getallasync.md)|取得先前所建立的所有繫結。|
|[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|藉由其識別碼取得指定的繫結。|
|[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|移除指定的繫結。|

## <a name="support-details"></a>支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此方法。空白儲存格表示 Office 主應用程式不支援此方法。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||||
|:-----|:-----|:-----|:-----|
||Office for Windows desktop|Office Online (在瀏覽器中)|Office for iPad|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## <a name="support-history"></a>支援歷程記錄



|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel 和 Word 的支援|
|1.1|對於 [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)、[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)和 [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)，新增支援在 Excel 增益集中繫結至矩陣資料做為表格繫結。|
|1.1|<ul><li>對於<a href="8fa0cb4a-fad1-4f2e-9a7e-5f7aa7789eca.htm">文件</a>屬性，新增在 Access 內容增益集中代表目前 Access 資料庫之 <span class="keyword">Document</span> 物件的存取權。</li><li>對於所有方法，新增支援在 Access 內容增益集中的表格繫結。 </li></ul>|
|1.0|已導入|
