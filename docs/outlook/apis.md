
# <a name="outlook-add-in-apis"></a>Outlook 增益集 API

若要在 Outlook 增益集中使用 API，您必須指定 Office.js 程式庫的位置、需求集合、結構描述和權限。

## <a name="office.js-library"></a>Office.js 程式庫

若要與 Outlook 增益集 API 互動，您必須使用 Office.js 中的 JavaScript API。CDN 的程式庫是 _https://appsforoffice.microsoft.com/lib/1/hosted/Office.js_。提交至 Office 市集的增益集必須藉由這個 CDN 參考 Office.js；它們不能使用本機的參考。 

在實作增益集 UI 的網頁 (.html、.aspx 或 .php 檔案) 的 **head** 標記中宣告 CDN，在 **script** 標記的 **src** 屬性中︰


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

當我們加入新的 API 時，Office.js 的 URL 會保持不變。只有當我們破壞現有 API 的行為時，我們才會變更 URL 中的版本。

> **重要事項：**在開發任何 Office 主應用程式的增益集時，請從頁面的 `<head>` 區段中參考適用於 Office 的 JavaScript API。如此一來，可確保在任何本文元素之前完全初始化 API。Office 主應用程式需要增益集在啟用的 5 秒內初始化。超過這個臨界值會使增益集宣告為沒有回應，並且會向使用者顯示錯誤訊息。  

## <a name="requirement-sets"></a>需求集合

所有 Outlook API 均屬於信箱需求集合。信箱需求集合具有版本，而我們發行的每一組新 API 都屬於較高版本的集合。在最新的 API 組合發行時並非所有 Outlook 用戶端都將可支援，但如果 Outlook 用戶端宣告支援某個需求集合，它就會支援該需求集合中的所有 API。 

若要控制增益集顯示的 Outlook 用戶端，請在資訊清單中指定最低需求集合版本。例如，如果您指定需求集合 1.3 版，增益集將不會顯示在任何不支援至少 1.3 的 Outlook 用戶端中。 

指定需求集合並不會將您的增益集限制在該版本的 API。如果增益集指定需求集合 1.1 版，但在支援 1.3 版的 Outlook 用戶端中執行時，增益集仍可使用 v1.3 API。需求集合只會控制要顯示增益集的 Outlook 用戶端。

若要從大於指定資訊清單中所指定的需求集合檢查任何 API 的可用性，您可以使用標準的 JavaScript：


```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> **附註：**對於已在資訊清單指定之需求集合版本中的任何 API 來說，開發人員不需要執行這類檢查。

指定支援您的案例重要 API 集合的最低需求集合，若無這些需求，增益集的重要功能就無法運作。您會在 **Requirements**、**Sets** 和 **Set** 元素的資訊清單中指定需求集。如需詳細資訊，請參閱 [Outlook 增益集資訊清單](../outlook/manifests/manifests.md)和[了解 Outlook API 需求集合](..\..\reference\outlook\tutorial-api-requirement-sets.md)。

**Methods** 元素不適用 Outlook 增益集，因此您無法宣告支援特定方法。


## <a name="permissions"></a>權限

增益集需要適當權限才能使用它所需的 API。有四個層級的權限。如需詳細資訊，請參閱[了解 Outlook 增益集的權限](../outlook/understanding-outlook-add-in-permissions.md)。


|**權限等級**|**描述**|
|:-----|:-----|
|受限|允許使用實體，而不是規則運算式。|
|讀取項目|除了 _Restricted_ 中允許的項目，它允許︰<ul><li>規則運算式</li><li>Outlook 增益集 API 讀取權限</li><li>取得項目屬性和回撥權杖</li></ul>|
|讀寫|除了 _Read item_ 中允許的項目，它允許︰<ul><li>除了 <b>makeEwsRequestAsync</b> 以外完整的 Outlook 增益集 API 存取</li><li>設定項目屬性</li></ul>|
|讀寫信箱|除了 _Read/write_ 中允許的項目，它允許︰<ul><li>建立、讀取、寫入項目及資料夾</li><li>傳送項目</li><li>呼叫 [makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md#makeewsrequestasyncdata-callback-usercontext)</li></ul>|
一般來說，您應該指定增益集所需的最小使用權限。權限是在資訊清單中的 **Permissions** 元素中宣告。如需詳細資訊，請參閱 [Outlook 增益集資訊清單](../outlook/manifests/manifests.md)。如需安全性問題的詳細資訊，請參閱 [Outlook 增益集的隱私權、權限和安全性](../outlook/../../docs/develop/privacy-and-security.md)。


## <a name="additional-resources"></a>其他資源

- [Outlook 增益集資訊清單](../outlook/manifests/manifests.md)

- [了解 Outlook API 需求集合](../../reference/outlook/tutorial-api-requirement-sets.md)
    
- [Outlook 增益集的隱私權、權限和安全性](../outlook/../../docs/develop/privacy-and-security.md)
    
