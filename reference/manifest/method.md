
# <a name="method-element"></a>Method 項目
指定適用於 Office 的 JavaScript API 中，為了啟用您的 Office 增益集所需的個別方法。

 **增益集類型︰**內容、工作窗格


## <a name="syntax:"></a>語法：


```XML
<Method Name="string "/>
```


## <a name="contained-in:"></a>內含於：

 _ [Methods](../../reference/manifest/methods.md)_


## <a name="attributes"></a>屬性



|**屬性**|**類型**|**必要**|**描述**|
|:-----|:-----|:-----|:-----|
|名稱|字串|必要|指定由其父物件限定的所需方法名稱。例如，若要指定 **getSelectedDataAsync** 方法，您必須指定 `"Document.getSelectedDataAsync"`。|

## <a name="remarks"></a>備註

郵件增益集不支援 **Methods** 和 **Method** 項目。如需有關需求集合的詳細資訊，請參閱[指定 Office 主機和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_intro)。


 >**重要**  因為沒有辦法指定個別方法的最小版本需求，為了確保方法可在執行階段使用，在您增益集的指令碼中呼叫方法時，您也應該使用 **if** 陳述式。如需有關如何執行這項操作的詳細資訊，請參閱 [了解適用於 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md#HostAPISupport_UsingIfStatements)。

