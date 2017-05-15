# <a name="rule-element"></a>Rule 元素

指定要針對此關聯式郵件增益集評估的啟用規則。

 **增益集類型：**郵件關聯式增益集

## <a name="contained-in"></a>內含於：

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md)

## <a name="attributes"></a>屬性

| 屬性 | 必要 | 描述 |
|:-----|:-----|:-----|
| **xsi:type** | 是 | 正在定義的規則類型。 |

規則類型可能是下列其中之一。

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch)
- [RuleCollection](#rulecollection)

## <a name="itemis-rule"></a>ItemIs 規則

定義規則，會評估如果選取的項目是指定的類型，則判斷值為 True。

### <a name="attributes"></a>屬性

| 屬性 | 必要 | 描述 |
|:-----|:-----|:-----|
| **ItemType** | 是 | 指定要比對的項目類型。可以是 `Message` 或 `Appointment`。`Message` 項目類型包括電子郵件、會議要求、會議回覆以及會議取消通知。 |
| **FormType** | 否 (在 [ExtensionPoint](extensionpoint.md) 內)，是 (在 [OfficeApp](officeapp.md) 內) | 指定應用程式是否會出現在項目的讀取或編輯表單中。可以是下列其中一項：`Read`、`Edit`、`ReadOrEdit`。 |
| **ItemClass** | 否 | 指定要比對的自訂郵件類別。如需詳細資訊，請參閱[啟動 Outlook 中特定郵件類別的郵件增益集](../../docs/outlook/manifests/activation-rules.md)。 |
| **IncludeSubClasses** | 否 | 指定如果項目是指定郵件類別的子類別，規則是否應評估為 True；預設值為 `false`。 |

### <a name="example"></a>範例

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a>ItemHasAttachment 規則

定義規則，會評估如果項目包含附件，則判斷值為 True。

### <a name="example"></a>範例

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity 規則

定義規則，會評估如果項目在其主旨或內文中，包含指定的實體類型文字，則判斷值為 True。

### <a name="attributes"></a>屬性

| 屬性 | 必要 | 描述 |
|:-----|:-----|:-----|
| **EntityType** | 是 | 指定規則若要評估為 True，所必須找到的實體類型。可以是下列其中一項：`MeetingSuggestion`、`TaskSuggestion`、`Address`、`Url`、`PhoneNumber`、`EmailAddress` 或 `Contact`。 |
| **RegExFilter** | 否 | 指定要針對此實體執行以啟用的規則運算式。 |
| **FilterName** | 否 | 指定規則運算式篩選的名稱，如此之還後可以在增益集的程式碼中參考此篩選。 |
| **IgnoreCase** | 否 | 指定在執行 **RegExFilter** 屬性指定的規則運算式時忽略大小寫。 |
| **Highlight** | 否 | **附註：**僅適用於 **ExtensionPoint** 元素內的 **Rule** 元素。指定用戶端應該如何反白顯示相符實體。可以是下列其中一項：`all`、`first`、`none`。若不指定，則預設值會設為 `all`。 |

### <a name="example"></a>範例

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch 規則

定義規則，會評估如果在項目的指定屬性中，可以找到相符的指定規則運算式，則判斷值為 True。

### <a name="attributes"></a>屬性

| 屬性 | 必要 | 描述 |
|:-----|:-----|:-----|
| **RegExName** | 是 | 指定規則運算式篩選的名稱，如此還可以在增益集的程式碼中參考此運算式。 |
| **RegExValue** | 是 | 指定要評估的規則運算式，評估後會決定是否應顯示郵件增益集。 |
| **PropertyName** | 是 | 指定規則運算式會評估的屬性名稱。可以是下列其中一項：`Subject`、`BodyAsPlaintext`、`BodyAsHtml` 或 `SenderSTMPAddress`。 |
| **IgnoreCase** | 否 | 指定在執行規則運算式時忽略大小寫。 |
| **Highlight** | 否 | **附註：**僅適用於 **ExtensionPoint** 元素內的 **Rule** 元素。指定用戶端應該如何反白顯示相符文字。可以是下列其中一項：`all`、`first`、`none`。若不指定，則預設值會設為 `all`。 |

### <a name="example"></a>範例

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

定義規則集合，和對其評估時會使用的邏輯運算子。

### <a name="attributes"></a>屬性

| 屬性 | 必要 | 描述 |
|:-----|:-----|:-----|
| **Mode** | 是 | 指定評估此規則集合時要使用的邏輯運算子。可以是 `And` 或 `Or`。 |

### <a name="example"></a>範例

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="additional-resources"></a>其他資源

- [啟動 Outlook 中特定郵件類別的郵件增益集](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx) 和 [Outlook 增益集的啟用規則](../../docs/outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins)
- [使 Outlook 項目中的字串與已知的實體相符](../../docs/outlook/match-strings-in-an-item-as-well-known-entities.md)    
- [使用規則運算式的啟用規則來顯示 Outlook 增益集](../../docs/outlook/use-regular-expressions-to-show-an-outlook-add-in.md)