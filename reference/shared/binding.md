
# Binding 物件
抽象類別，代表繫結至文件的某個區段。

|||
|:-----|:-----|
|**主機︰**|Access、Excel、Word|
|**可用於[需求集合](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBinding, TableBinding, TextBinding|
|**上次變更於 TableBinding**|1.1|

```js
Office.context.document.bindings.getByIdAsync(id);
```

## 成員


**物件**


|**名稱**|**說明**|
|:-----|:-----|
|[MatrixBinding](../../reference/shared/binding.matrixbinding.md)|表示資料列和資料行的兩個維度中的繫結。|
|[TableBinding](../../reference/shared/binding.tablebinding.md)|表示資料列和資料行的兩個維度中的繫結，選擇性地包含標頭。|
|[TextBinding](../../reference/shared/binding.textbinding.md)|表示文件中的繫結文字選取。|

**屬性**


|**名稱**|**說明**|
|:-----|:-----|
|[文件](../../reference/shared/binding.document.md)|取得與繫結相關聯的 **Document** 物件。|
|[id](../../reference/shared/binding.id.md)|取得物件的識別碼。|
|[類型](../../reference/shared/binding.type.md)|取得繫結的類型。|

**方法**


|**名稱**|**說明**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md)|加入指定的事件類型繫結的處理常式。|
|[getDataAsync](../../reference/shared/binding.getdataasync.md)|傳回包含在繫結的資料。|
|[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|移除指定的事件類型繫結中所指定的處理常式。|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|將資料寫入指定的繫結物件所代表文件的繫結區段。|
|[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|設定或更新指定項目和繫結表格中資料的格式設定。|

**事件**


|**名稱**|**說明**|
|:-----|:-----|
|[bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md)|繫結內的資料變更時，就會發生。|
|[bindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md)|繫結內的選取項目變更時，就會發生。|

## 備註

**Binding** 物件公開所有繫結擁有的功能，無論類型為何。

**Binding** 物件永遠不會直接呼叫。它是代表每一種繫結之物件的抽象父系類別：[MatrixBinding](../../reference/shared/binding.matrixbinding.md)、[TableBinding](../../reference/shared/binding.tablebinding.md) 或 [TextBinding](../../reference/shared/binding.textbinding.md)。這三種物件皆從 **Binding** 物件繼承 **getDataAsync** 和 **setDataAsync** 方法，可讓您與繫結中的資料互動。它們也繼承 **id** 和 **type** 屬性，以查詢這些屬性值。此外，**MatrixBinding** 與 **TableBinding** 物件公開專用於矩陣與表格功能的其他方法，例如計算列數與欄數。


## 支援詳細資料


支援 **Binding** 物件的每個 API 成員在 Office 主應用程式之間有所不同。如需瞭解主機支援資訊，請參閱每個成員主題的「支援詳細資料」一節。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||
|:-----|:-----|
|**可用於需求集合**|MatrixBinding, TableBinding, TextBinding|
|**增益集類型**|內容、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|
