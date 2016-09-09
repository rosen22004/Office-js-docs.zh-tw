
# Rule 項目
指定要針對此郵件增益集評估的啟用規則。

 **增益集類型︰**郵件


## 語法：

 **ItemIs 規則** - 定義規則，會評估如果選取的項目是指定的類型，則判斷值為 True。


```XML
<Rule xsi:type="ItemIs" 
   ItemType= ["Appointment" | "Message"]
   FormType=["Read" | "Edit" | "ReadOrEdit"] 
   ItemClass = "string " 
   IncludeSubClasses=["true" | "false"] />
```

 **ItemHasAttachment 規則** - 定義規則，會評估如果項目包含附件，則判斷值為 True。




```XML
<Rule xsi:type="ItemHasAttachment"  />
```

 **ItemHasKnownEntity** - 定義規則，會評估如果項目在其主旨或內文中，包含指定的實體類型文字，則判斷值為 True。




```XML
<Rule xsi:type="ItemHasKnownEntity" 
  EntityType=["MeetingSuggestion" | "TaskSuggestion" |"Address" | "Url" | "PhoneNumber" | "EmailAddress" | "Contact" ]
  RegExFilter="string "
  FilterName="string "
  IgnoreCase=["true | false"]/>
```

 **ItemHasRegularExpressionMatch 規則** - 定義規則，會評估如果在項目的指定屬性中，可以找到相符的指定規則運算式，則判斷值為 True。




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="string " 
    RegExValue="string " 
    PropertyName=["Subject" | "BodyAsPlaintext" | "BodyAsHtml" | "SenderSTMPAddress"]
    IgnoreCase=["true" | "false"]
/>
```

 **RuleCollection 規則** - 定義規則集合，和對其評估時會使用的邏輯運算子。




```XML
<Rule xsi:type="RuleCollection" Mode=["And" | "Or"]>
   ...
</Rule>
```


## 內含於：

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## 屬性：

 **ItemIs 規則屬性**



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|ItemType|ItemType (字串)|必要|指定要比對的項目類型。可以是下列其中一項：

|**ItemType**|**對應的 ItemClass**|
|:-----|:-----|
|約會|IPM.Appointment|
|Message(1)|包括電子郵件訊息、會議邀請、回覆和取消。|
|
|FormType|ItemFormType (字串)|必要|指定應用程式是否會出現在項目的讀取或編輯表單中。 可以是下列其中一項。|

|**FormType**|**說明**|
|:-----|:-----|
|讀取|指定只有在 (指定的 **ItemType**) 讀取表單中，啟動郵件增益集。|
|Edit|指定只有在 (指定的 **ItemType**) 撰寫表單中，啟動郵件增益集。|
|ReadOrEdit|指定在 (指定的 **ItemType**) 讀取和撰寫表單中，都啟動郵件增益集。|
|ItemClass|String|選用|指定要比對的自訂郵件類別。如需詳細資訊，請參閱[啟動 Outlook 中特定郵件類別的郵件增益集](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx)。|
|IncludeSubClasses|布林值|選用|指定如果項目是指定郵件類別的子類別，規則是否應評估為 True；預設值為 False。|


(1) 以下是對應的郵件類別︰IPM.NoteIPM.Schedule.Meeting.RequestIPM.Schedule.Meeting.NegIPM.Schedule.Meeting.PosIPM.Schedule.Meeting.TentIPM.Schedule.Meeting.Canceled。

 **ItemHasAttachment 規則屬性**

無。

 **ItemHasKnownEntity 規則屬性**



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|EntityType|KnownEntityType (字串)|必要|指定規則若要評估為 True，所必須找到的實體類型。可以是下列其中一項：

|**KnownEntityType**|**說明**|
|:-----|:-----|
|MeetingSuggestion|參考事件或會議，且根據樣式辨識所識別的文字。|
|TaskSuggestion| 包含可採取動作的片語，且根據樣式辨識所識別的文字。|
|Address|參考美國的郵寄地址，且根據樣式辨識所識別的文字。|
|URL|包含檔案名稱或網站位址 URL，且根據樣式辨識所識別的文字。|
|PhoneNumber| 根據模式辨識識別為北美電話號碼的一串數字。|
|EmailAddress|包含 SMTP 格式電子郵件地址，且根據樣式辨識所識別的文字。|
|Contact|包含連絡人資訊，且根據樣式辨識所識別的文字。|
|RegExFilter|String|選用|指定要針對此實體執行以啟用的規則運算式。|
|FilterName|String|選用|指定規則運算式篩選的名稱，如此之還後可以在增益集的程式碼中參考此篩選。|
|IgnoreCase|布林值|選用|指定在執行 **RegExFilter** 屬性指定的規則運算式時忽略大小寫。|
 **ItemHasRegularExpressionMatch 規則屬性**



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|RegExName|String|必要|指定規則運算式篩選的名稱，如此還可以在增益集的程式碼中參考此運算式。|
|RegExValue|String|必要|指定要評估的規則運算式，評估後會決定是否應顯示郵件增益集。 |
|PropertyName|PropertyName (字串)|必要|指定規則運算式會評估的屬性名稱。可以是下列其中一項：

|**PropertyName**|**說明**|
|:-----|:-----|
|主旨|對項目主旨評估規則運算式。|
|BodyAsPlaintext|針對項目主旨評估規則運算式。|
|BodyAsHtml|如果能以 HTML 呈現項目內文，則針對內文評估規則運算式。|
|SenderSTMPAddress|針對項目寄件人的 SMTP 位址評估規則運算式。|
|IgnoreCase|布林值|選用|指定在執行規則運算式時忽略大小寫。|
 **RuleCollection 規則屬性**



|**屬性**|**類型**|**必要**|**說明**|
|:-----|:-----|:-----|:-----|
|Mode|string|必要|指定評估此規則集合時要使用的邏輯運算子。可以是︰"And" 或 "Or"。|

## 其他資源



- 
  [啟動 Outlook 中特定郵件類別的郵件增益集](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx) 和 [Outlook 增益集的啟用規則](../../docs/outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins)
    
- [使 Outlook 項目中的字串與已知的實體相符](../../docs/outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [使用規則運算式的啟用規則來顯示 Outlook 增益集](../../docs/outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
