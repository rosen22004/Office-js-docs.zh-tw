# <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API 需求集合

需求集合是 API 成員的具名群組。Office 增益集使用資訊清單中所指定的需求集合，或使用執行階段檢查，以判定 Office 主應用程式是否支援增益集所需的 API。如需詳細資訊，請參閱[指定 Office 主應用程式及 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

下表列出 OneNote 需求集、支援需求集的 Office 主應用程式，以及組建版本或可用性日期。

|  需求集合  |  Office Online | 
|:-----|:-----|
| OneNoteApi 1.1  | 2016 年 9 月 |  

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 需求集合
如需通用 API 需求集合的詳細資訊，請參閱 [Office 通用 API 需求集合](office-add-in-requirement-sets.md)。

## <a name="onenote-javascript-api-11"></a>OneNote JavaScript API 1.1 
OneNote JavaScript API 1.1 是 API 的第一個版本。如需 API 的詳細資訊，請參閱 [OneNote JavaScript API](../../docs/onenote/onenote-add-ins-programming-overview.md) 參考主題。

## <a name="runtime-requirement-support-check"></a>執行階段需求支援檢查

在執行階段期間，增益集可以藉由執行下列檢查來檢查特定主機是否支援 API 需求集合： 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a>基於資訊清單的需求支援檢查

在增益集資訊清單中使用 Requirements 元素來指定增益集必須使用的關鍵需求集合或 API 成員。如果 Office 主機或平台不支援 Requirements 元素中指定的需求集合或 API 成員，增益集將不會在該主機或平台上執行，且不會顯示在我的增益集中。

下列程式碼範例會顯示支援 OneNoteApi 需求集合 1.1 版的所有 Office 主應用程式中載入的增益集。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```



## <a name="additional-resources"></a>其他資源

- [指定 Office 主應用程式和 API 需求](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office 增益集的 XML 資訊清單](../../docs/overview/add-in-manifests.md)
