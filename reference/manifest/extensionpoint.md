# <a name="extensionpoint-element"></a>ExtensionPoint 元素

 在 Office UI 中定義增益集公開功能的位置。**ExtensionPoint** 元素是 [DesktopFormFactor](./desktopformfactor.md) 或 [MobileFormFactor](./mobileformfactor.md) 的子元素。 

## <a name="attributes"></a>屬性

|  屬性  |  必要  |  描述  |
|:-----|:-----|:-----|
|  **xsi:type**  |  是  | 所定義的擴充點型別。|


## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Word、Excel、PowerPoint 及 OneNote 增益集命令的擴充點

- **PrimaryCommandSurface** - Office 中的功能區。
- **ContextMenu** - 當您以滑鼠右鍵按一下 Office UI 時所顯示的快顯功能表。

下列範例顯示如何將 **ExtensionPoint** 元素與 **PrimaryCommandSurface** 和 **ContextMenu** 屬性值搭配使用，以及應使用的每一個子元素。


 >**重要**  如需包含 ID 屬性的元素，請確定提供唯一的 ID。我們建議您使用貴公司的名稱以及您的 ID。例如，您可以使用下列格式。<CustomTab id="mycompanyname.mygroupname">


```XML
 <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Contoso Tab">
            <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
             <!-- <OfficeTab id="TabData"> -->
              <Label resid="residLabel4" />
              <Group id="Group1Id12">
                <Label resid="residLabel4" />
                <Icon>
                  <bt:Image size="16" resid="icon1_32x32" />
                  <bt:Image size="32" resid="icon1_32x32" />
                  <bt:Image size="80" resid="icon1_32x32" />
                </Icon>
                <Tooltip resid="residToolTip" />
                <Control xsi:type="Button" id="Button1Id1">

                   <!-- information about the control -->
                </Control>
                <!-- other controls, as needed -->
              </Group>
            </CustomTab>
          </ExtensionPoint>

        <ExtensionPoint xsi:type="ContextMenu">
          <OfficeMenu id="ContextMenuCell">
            <Control xsi:type="Menu" id="ContextMenu2">
                   <!-- information about the control -->
            </Control>
           <!-- other controls, as needed -->
          </OfficeMenu>
         </ExtensionPoint>
```

**子元素**
 
|** 元素**|**描述**|
|:-----|:-----|
|**CustomTab**|如果您想要將自訂索引標籤加入至功能區 (使用 **PrimaryCommandSurface**)，則為必要。如果您使用 **CustomTab** 元素，您無法使用 **OfficeTab** 元素。**id** 屬性為必要項目。|
|**OfficeTab**|如果您想要擴充預設 Office 功能區索引標籤 (使用 **PrimaryCommandSurface**)，則為必要。如果您使用 **OfficeTab** 元素，您無法使用 **CustomTab** 元素。如需詳細資訊，請參閱 [OfficeTab](officetab.md)。|
|**OfficeMenu**|如果要將增益集命令加入至預設的快顯功能表 (使用 **ContextMenu**)，則為必要。**id** 屬性必須設定為︰ <br/> - **ContextMenuText** (適用於 Excel 或 Word)。選取文字，然後使用者在選取的文字上按一下滑鼠右鍵時，會顯示快顯功能表上的項目。 <br/>Excel 的  - **ContextMenuCell**。當使用者在試算表的儲存格上按一下滑鼠右鍵，會在快顯功能表上顯示項目。|
|**群組**|索引標籤上的一群使用者介面擴充點。一個群組可以有最多六個控制項。**id** 屬性為必要項目。它是最多為 125 個字元的字串。|
|**標籤**|必要。群組的標籤。**resid** 屬性必須設為 **String** 元素的 **id** 屬性值。**String** 元素是 **ShortStrings** 元素的子元素，為 **Resources** 元素的子元素。|
|**圖示**|必要。指定要用於小型裝置上，或顯示太多按鈕時所使用的群組圖示。**resid** 屬性必須設為 **Image** 元素的 **id** 屬性值。**Image** 元素是 **Images** 元素的子元素，為 **Resources** 元素的子元素。**size** 屬性會提供影像的大小，單位為像素。需要三個影像大小：16、32 和 80。也支援五個選擇性大小︰20、24、40、48 及 64。|
|**Tooltip**|選用。群組的工具提示。**resid** 屬性必須設為 **String** 元素的 **id** 屬性值。**String** 元素是 **LongStrings** 元素的子元素，為 **Resources** 元素的子元素。|
|**Control**|每個群組至少需要一個控制項。**Control** 元素可以是 **Button** 或 **Menu**。使用 **Menu** 指定按鈕控制項的下拉式清單。目前僅支援按鈕和功能表。請參閱[按鈕控制項](#button-controls)和[功能表控制項](#menu-controls)小節以取得詳細資訊。<br/>**附註**  若要使疑難排解更加容易，我們建議一次加入一個 **Control** 元素和相關的 **Resources** 子元素。

## <a name="extension-points-for-outlook-add-in-commands"></a>Outlook 增益集命令的擴充點

- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (只能在 [DesktopFormFactor](./desktopformfactor.md) 中使用。)
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [事件](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface
這個擴充點會將按鈕放在郵件讀取檢視的命令介面中。在 Outlook 桌面中，這將會出現在功能區。

**子元素**

|  元素 |  描述  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

#### <a name="officetab-example"></a>OfficeTab 範例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 範例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface
這個擴充點會使用郵件撰寫表單，將增益集的按鈕放在功能區上。 

**子元素**

|  元素 |  描述  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

#### <a name="officetab-example"></a>OfficeTab 範例
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 範例

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

這個擴充點會將向會議召集人顯示的表單按鈕放在功能區中。 

**子元素**

|  元素 |  描述  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

#### <a name="officetab-example"></a>OfficeTab 範例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 範例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

這個擴充點會將向會議出席者顯示的表單按鈕放在功能區中。 

**子元素**

|  元素 |  描述  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

#### <a name="officetab-example"></a>OfficeTab 範例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 範例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>模組

這個擴充點會將按鈕放在模組擴充的功能區上。 

**子元素**

|  元素 |  描述  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface
這個擴充點會將按鈕放在行動裝置外觀尺寸中郵件讀取檢視的命令介面中。

> **附註：**只有 iOS 版 Outlook 支援此元素類型。

**子元素**

|  元素 |  描述  |
|:-----|:-----|
|  [Group](./group.md) |  將一組按鈕新增至命令介面。  |
|  [Control](./control.md) |  將單一按鈕新增至命令介面。  |

此類型的 **ExtensionPoint** 元素只可以有一個子元素 (**Group** 元素或 **Control** 元素)。

這個擴充點內涵的 **Control** 元素必須將 **xsi:type** 屬性設定為 `MobileButton`。

#### <a name="group-example"></a>Group 範例
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
    <!-- one or more Control elements -->
  </Group>
</ExtensionPoint>
```

#### <a name="control-example"></a>Control 範例
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Control id="mobileButton1" xsi:type="MobileButton">
    <!-- Control definition -->
  </Control>
</ExtensionPoint>
```

### <a name="events"></a>事件
此擴充點會新增特定事件的事件處理常式。

> **附註：**僅 Office 365 中的 Outlook 網頁版支援此元素類型。

|  元素 |  描述  |
|:-----|:-----|
|  [事件](./event.md) |  指定事件與事件處理常式的函數。  |

#### <a name="itemsend-event-example"></a>ItemSend 事件範例
```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a>DetectedEntity
這個擴充點會在指定實體類型上新增關聯式增益集啟動。

抑制 [VersionOverrides](./versionoverrides.md) 元素必須具備 `VersionOverridesV1_1` 的 `xsi:type` 屬性值。

> **附註：**僅 Office 365 中的 Outlook 網頁版支援此元素類型。

|  元素 |  描述  |
|:-----|:-----|
|  [Label](#label) |  在關聯式視窗中指定增益集的標籤。  |
|  [SourceLocation](./sourcelocation.md) |  指定關聯式視窗的 URL。  |
|  [Rule](./rule.md) |  指定規則，該規則決定增益集啟動的時機。  |

#### <a name="label"></a>標籤

必要。群組的標籤。**resid** 屬性必須設定為  **Resources** 元素的 **ShortStrings** 元素中，[String](./resources.md#shortstrings) 元素的 [id](./resources.md) 屬性值。

#### <a name="highlight-requirements"></a>反白顯示需求

使用者啟動關聯式增益集的唯一方式是與反白顯示的實體互動。開發人員可以控制要反白顯示哪些實體，方法是使用 `ItemHasKnownEntity` 和 `ItemHasRegularExpressionMatch` 規則類型之 `Rule` 元素的 `Highlight`屬性。

但是，要注意有某些限制。這些限制是為了確保適用的訊息或約會中永遠都有反白顯示的實體，讓使用者能夠啟動增益集。

- `EmailAddress` 和 `Url` 實體類型無法反白顯示，因此無法用來啟動增益集。
- 如果使用單一規則，`Highlight` 必須設定為 `all` 或 `first`。
- 如果使用 `RuleCollection` 規則類型與 `Mode="AND"` 以結合多個規則類型，至少其中一個規則「必須」將 `Highlight` 設定為 `all` 或 `first`。
- 如果使用 `RuleCollection` 規則類型與 `Mode="OR"` 以結合多個規則類型，所有規則「必須」將 `Highlight` 設定為 `all` 或 `first`。

#### <a name="detectedentity-event-example"></a>DetectedEntity 事件範例
```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```