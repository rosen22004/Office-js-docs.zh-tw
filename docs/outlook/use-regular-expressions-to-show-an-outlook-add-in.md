
# 使用規則運算式的啟用規則來顯示 Outlook 增益集

您可以指定規則運算式規則讓 Outlook 增益集在讀取案例中啟動 - 當使用者在讀取窗格或檢查程式中檢視郵件或約會時，Outlook 會評估規則運算式規則，來決定它是否應該啟動增益集。使用者在撰寫項目時，Outlook 不會評估這些規則。另外還有其他 Outlook 不會啟動增益集的案例，例如，受資訊版權管理 (IRM) 保護或在 [垃圾郵件] 資料夾中的項目。如需詳細資訊，請參閱 [Outlook 增益集的啟用規則](../outlook/manifests/activation-rules.md)。

您可以指定規則運算式做為 [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) 規則或 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 增益集 XML 資訊清單中規則的一部分。Outlook 會根據用戶端電腦上瀏覽器所使用的 JavaScript 直譯器的規則來評估規則運算式。Outlook 支援所有 XML 處理器也都支援的相同特殊字元清單。下表列出這些特殊字元。您可以在規則運算式中使用這些字元，方法為指定對應字元的逸出序列，如下表所述。



|**字元**|**說明**|**要使用的逸出序列**|
|:-----|:-----|:-----|
|"|雙引號|&amp;quot;|
|&amp;|& 符號|&amp;amp;|
|'|' 單引號|&amp;apos;|
|<|小於符號|&amp;lt;|
|>|大於符號|&amp;gt;|

## ItemHasRegularExpressionMatch 規則


根據受支援屬性的特定值，**ItemHasRegularExpressionMatch** 規則對於控制增益集啟動很有幫助。**ItemHasRegularExpressionMatch** 規則具有下列屬性。



|**屬性名稱**|**說明**|
|:-----|:-----|
|**RegExName**|指定規則運算式篩選的名稱，如此還可以在增益集的程式碼中參考此運算式。|
|**RegExValue**|指定要評估的規則運算式，評估後會決定是否應顯示增益集。|
|**PropertyName**|指定規則運算式會評估的屬性名稱。允許的值為 **BodyAsHTML**、**BodyAsPlaintext**、**SenderSMTPAddress** 和 **Subject**。如果您指定 **BodyAsHTML**，則只要項目本文是 HTML，Outlook 會套用規則運算式，否則 Outlook 不會傳回該規則運算式任何符項目。因為約會永遠會以 RTF 格式儲存，指定 **BodyAsHTML** 的規則運算式不會符合約會項目本文中的任何字串。如果您指定 **BodyAsPlaintext**，則 Outlook 會一律在項目本文上套用規則運算式。|
|**IgnoreCase**|指定當比對由 **RegExName**所指定的規則運算式時，是否要忽略大小寫。|

### 在規則中使用規則運算式的最佳做法

當您使用規則運算式時，請特別注意下列項目︰


- 如果您在項目的本文屬性上指定 **ItemHasRegularExpressionMatch** 規則，規則運算式應該進一步篩選本文，且不應該嘗試傳回項目的整個本文。使用規則運算式 (例如 `.*`) 來嘗試取得項目的整個本文並不會永遠傳回預期的結果。
    
- 一個瀏覽器上傳回的純文字本文可能會與另一個瀏覽器上的有細微差異。如果您使用 [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) 規則與 **BodyAsPlaintext** 做為 **PropertyName** 屬性，請測試您的增益集支援的所有瀏覽器上的規則運算式。
    
    由於不同瀏覽器會使用不同的方式，以取得所選項目的文字內文，請務必確定您的規則運算式支援這些內文文字間可能出現的細微差別。 例如，某些瀏覽器 (例如 Internet Explorer 9) 使用 DOM 屬性 **innerText**，而其他瀏覽器 (例如 Firefox) 使用 **.textContent()** 方法取得項目文字內文。 此外，不同的瀏覽器可能會以不同方式傳回換行︰Internet Explorer 用 "\r\n"，而 Firefox 和 Chrome 用 "\n"。 如需詳細資訊，請參閱 [W3C DOM 相容性 - HTML](http://www.quirksmode.org/dom/w3c_html.mdl#t07)。
    
- 項目的 HTML 本文在 Outlook 豐富型用戶端與 Outlook Web App 或裝置用 OWA 之間稍有不同。仔細定義您的規則運算式。舉例來說，請考慮下列使用於 **ItemHasRegularExpressionMatch** 規則與 **BodyAsHTML** 中做為 **PropertyName** 屬性值的規則運算式：
    
```
      http.*\.contoso\.com
```


    A rule with this regular expression would match the string "http-equiv="Content-Type" which exists in the HTML body of an item in an Outlook rich client, as part of the following  **META** tag:
    

```HTML
      <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii">
```


相同的規則不會在 Outlook Web App 和裝置用 OWA 中傳回此相符項目，因為這些主機中的 HTML 本文不包含該 **META** 標記。這可能會影響增益集是否已針對各種 Outlook 用戶端適當地啟用。在這個範例中，請改用下列規則運算式︰
    

```
      http://.*\.contoso\.com/
```

- 依據套用規則運算式的主應用程式、裝置類型或屬性，當設計規則運算式做為啟用規則時，您需要注意每一個主機的其他最佳做法與限制。如需詳細資訊，請參閱[適用於 Outlook 增益集的 JavaScript API 和啟動的限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    

### 範例

只要寄件者的 SMTP 電子郵件地址符合「@contoso」，無論大寫或小寫字元，下列 **ItemHasRegularExpressionMatch** 規則就會啟用增益集。


```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]" 
    PropertyName="SenderSMTPAddress"
/>
```

下面是使用 **IgnoreCase** 屬性來指定相同規則運算式的另一種方式。




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@contoso" 
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

只要目前項目的本文中包含某個股票代號，下列 **ItemHasRegularExpressionMatch** 規則就會啟用增益集。




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    PropertyName="BodyAsPlaintext" 
    RegExName="TickerSymbols" 
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```


## ItemHasKnownEntity 規則



  **ItemHasKnownEntity** 規則會根據選取項目的主旨或本文中的實體是否存在來啟用增益集。[KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) 類型會定義支援的實體。在 **ItemHasKnownEntity** 規則上套用規則運算式會提供便利性，其中啟用是根據實體 (例如，特定的 URL 集，或具有特定區碼的電話號碼) 子集的值而定。


 >
  **附註**  Outlook 僅可以英文擷取實體字串，無論資訊清單中指定的預設地區設定。僅郵件但非約會支援 **MeetingSuggestion** 實體類型。您無法從 [寄件備份] 資料夾中的項目擷取實體，也無法使用 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 規則啟動 [寄件備份] 資料夾中項目的增益集。

**ItemHasKnownEntity** 規則支援下表中的屬性。請注意，雖然在 **ItemHasKnownEntity** 規則中指定規則運算式是選擇性的，如果您選擇使用規則運算式做為實體篩選條件，必須指定 **RegExFilter** 和 **FilterName** 屬性。



|**屬性名稱**|**說明**|
|:-----|:-----|
|**EntityType**|指定規則若要評估為 **true**，所必須找到的實體類型。使用多個規則來指定多個實體類型。|
|**RegExFilter**|指定規則運算式，它會進一步篩選由 **EntityType** 所指定之實體的執行個體。|
|**FilterName**|指定由 **RegExFilter** 所指定的規則運算式篩選的名稱，如此之後還可以由程式碼參考此篩選。|
|**IgnoreCase**|指定當比對由 **RegExFilter** 所指定的規則運算式時，是否要忽略大小寫。|

### 範例

每當目前項目的主旨或本文中有 URL，且 URL 包含字串「youtube」(無論字串的大小寫) 時，下列 **ItemHasKnownEntity** 規則會啟動增益集。


```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```


## 在程式碼中使用規則運算式的結果


您可以在目前的項目上使用下列方法取得規則運算式的相符項目︰


- 針對所有在增益集的 [ItemHasRegularExpressionMatch](../../reference/outlook/Office.context.mailbox.item.md) 和 **ItemHasKnownEntity** 規則中指定的規則運算式，**getRegExMatches** 會傳回目前項目中的相符項目。
    
- 針對在增益集的 [ItemHasRegularExpressionMatch](../../reference/outlook/Office.context.mailbox.item.md) 規則中指定的已識別規則運算式，**getRegExMatchesByName** 會傳回目前項目中的相符項目。
    
- [getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md) 傳回整個實體的執行個體，其包含增益集的 **ItemHasKnownEntity** 規則中指定的已識別規則運算式的相符項目。
    
當評估規則運算式時，相符項目會傳回至陣列物件中的增益集。針對 **getRegExMatches**，該物件具有規則運算式的名稱的識別項。 


 >**附註**  Outlook 豐富型用戶端不會在陣列中以任何特定順序傳回相符項目。此外，您不應該假設 Outlook 豐富型用戶端在此陣列中會以與 Outlook Web App 或裝置用 OWA 相同的順序傳回相符項目，即使當您在相同信箱中的相同項目的每個用戶端上執行相同的增益集。如需了解在 Outlook 豐富型用戶端與 Outlook Web App 或裝置用 OWA 之間處理規則運算式的其他差異，請參閱[適用於 Outlook 增益集的 JavaScript API 和啟動的限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)。


### 範例

下列是規則集合的範例，其中包含 **ItemHasRegularExpressionMatch** 規則與名為 `videoURL`的規則運算式。


```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="VideoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="Body"/>
</Rule>
```

下列範例會使用目前項目的 **getRegExMatches** 將變數 `videos` 設定為前述 **ItemHasRegularExpressionMatch** 規則的結果。




```
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

多個相符項目會在該物件中儲存為陣列元素。下列程式碼範例會示範如何反覆查看名為 `reg1` 的規則運算式的相符項目，以建立要顯示成 HTML 格式的字串。




```js
function initDialer() 
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = _Item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }
    myCell.innerHTML = myString;
}

```

下列是 **ItemHasKnownEntity** 規則的範例，指定 **MeetingSuggestion** 實體以及名為 `CampSuggestion` 的規則運算式。如果 Outlook 偵測到目前選取的項目包含會議建議，且主旨或本文包含「WonderCamp」一詞，則會啟動增益集。




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

下列程式碼範例會使用目前項目的 **getFilteredEntitiesByName(name)** 來設定變數 `suggestions` 以取得與前述 **ItemHasKnownEntity** 規則偵測到的會議建議陣列。




```
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName(CampSuggestion);
```


## 其他資源



- [建立讀取格式的 Outlook 增益集](../outlook/read-scenario.md)
    
- [Outlook 增益集的啟用規則](../outlook/manifests/activation-rules.md)
    
- [適用於 Outlook 增益集的 JavaScript API 和啟動的限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [使 Outlook 項目中的字串與已知的實體相符](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [在 .NET Framework 中使用規則運算式的最佳做法](http://msdn.microsoft.com/en-us/library/618e5afb-3a97-440d-831a-70e4c526a51c%28Office.15%29.aspx)
    
