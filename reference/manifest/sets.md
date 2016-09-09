
# Sets 項目
指定適用於 Office 的 JavaScript API 中，為了啟用您的 Office 增益集所需的最小子集合。

 **增益集類型︰**內容、工作窗格、郵件


## 語法：


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## 內含於：

[需求](../../reference/manifest/requirements.md)


## 可以包含︰

[Set](../../reference/manifest/set.md)


## 屬性



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|String|選用|指定所有子 **Set** 項目預設的 [MinVersion](../../reference/manifest/set.md) 屬性值。預設值為「1.1」。|

## 備註

如需有關需求集合的詳細資訊，請參閱[指定 Office 主機和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

如需有關 **Set** 項目的 **MinVersion** 屬性和 **Sets** 項目的 **DefaultMinVersion** 屬性的詳細資訊，請參閱[設定資訊清單中的 Requirements 項目](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。

