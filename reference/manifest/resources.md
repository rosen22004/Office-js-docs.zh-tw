# Resources 元素

包含圖示、字串，和 [VersionOverrides](./versionoverrides.md) 節點的 URL。 資訊清單元素可使用資源的 **id** 來指定資源。 這可協助保護可管理的資訊清單大小，特別是在資源具有不同地區設定的版本時。 **id** 在資訊清單內必須唯一，且最多有 32 個字元。

每個資源都可以具有一或多個 **Override** 子元素，為特定地區設定定義不同的資源。

## 子元素

|  元素 |  類型	  |  說明  |
|:-----|:-----|:-----|
|  [影像](#影像)            |  Image   |  提供圖示之影像的 HTTPS URL。 |
|  **URL**                |  URL     |  提供 HTTPS URL 位置。 URL 最多可以有 2048 個字元。 |
|  **ShortStrings** |  字串  |  **Label** 和 **Title** 元素的文字。 每個 **String** 包含最多 125 個字元。|
|  **LongStrings**  |  字串  | **Description** 屬性的文字。 每個 **String** 包含最多 250 個字元。|

>**附註**  您必須在 **Image** 和 **Url** 元素中，為所有 URL 使用安全通訊端層 (SSL)。

### 影像
每個圖示必須要有三個 **Images** 元素，三種必要大小各一個︰
- 16x16
- 32x32
- 80x80

也支援下列其他的大小，但非必要︰
- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> **重要事項：** Outlook 需要針對效能目的的快取影像資源的能力。 基於這個理由，裝載影像資源的伺服器不能將任何 CACHE-CONTROL 指示詞新增至回應標頭。 這會導致 Outlook 自動取代一般或預設影像。    


## 資源範例 

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/Images/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/Images/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/Images/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```


```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```

