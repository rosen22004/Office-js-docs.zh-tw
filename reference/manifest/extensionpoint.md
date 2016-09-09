# ExtensionPoint 元素

 在 Office UI 中定義增益集公開功能的位置。 **ExtensionPoint** 元素是 [FormFactor](./formfactor.md) 的子元素。 

## 屬性

|  屬性  |  必要  |  說明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  是  | 所定義的擴充點型別。|


## Word、Excel、PowerPoint 及 OneNote 增益集命令的擴充點

- **PrimaryCommandSurface** - Office 中的功能區。
- **ContextMenu** - 當您以滑鼠右鍵按一下 Office UI 時所顯示的快顯功能表。

下列範例顯示如何將 **ExtensionPoint** 元素與 **PrimaryCommandSurface** 和 **ContextMenu** 屬性值搭配使用，以及應使用的每一個子元素。


 >**重要**  如需包含 ID 屬性的元素，請確定提供唯一的 ID。 我們建議您使用貴公司的名稱以及您的 ID。 例如，您可以使用下列格式。<CustomTab id="mycompanyname.mygroupname">


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
 
|**元素**|**說明**|
|:-----|:-----|
|**CustomTab**|如果您想要將自訂索引標籤加入至功能區 (使用 **PrimaryCommandSurface**)，則為必要。如果您使用 **CustomTab** 元素，您無法使用 **OfficeTab** 元素。**id** 屬性為必要項目。|
|**OfficeTab**|如果您想要擴充預設 Office 功能區索引標籤 (使用 **PrimaryCommandSurface**)，則為必要。 如果您使用 **OfficeTab** 元素，您無法使用 **CustomTab** 元素。 如需詳細資訊，請參閱 [OfficeTab](officetab.md)。|
|**OfficeMenu**|如果要將增益集命令加入至預設的快顯功能表 (使用 **ContextMenu**)，則為必要。 **id** 屬性必須設定為︰ <br/>Excel 或 Word 的  - **ContextMenuText**。 選取文字，然後使用者在選取的文字上按一下滑鼠右鍵時，會顯示快顯功能表上的項目。 <br/>Excel 的  - **ContextMenuCell**。 當使用者在試算表的儲存格上按一下滑鼠右鍵，會在快顯功能表上顯示項目。|
|**群組**|索引標籤上的一群使用者介面擴充點。一個群組可以有最多六個控制項。**id** 屬性為必要項目。它是最多為 125 個字元的字串。|
|**標籤**|必要。 群組的標籤。 **resid** 屬性必須設為 **String** 元素的 **id** 屬性值。 **String** 元素是 **ShortStrings** 元素的子元素，為 **Resources** 元素的子元素。|
|**圖示**|必要。 指定要用於小型裝置上，或顯示太多按鈕時所使用的群組圖示。 **resid** 屬性必須設為 **Image** 元素的 **id** 屬性值。 **Image** 元素是 **Images** 元素的子元素，為 **Resources** 元素的子元素。 **size** 屬性會提供影像的大小，單位為像素。 需要三個影像大小：16、32 和 80。 也支援五個選擇性大小︰20、24、40、48 及 64。|
|**工具提示**|選用。 群組的工具提示。 **resid** 屬性必須設為 **String** 元素的 **id** 屬性值。 **String** 元素是 **LongStrings** 元素的子元素，為 **Resources** 元素的子元素。|
|**控制項**|每個群組至少需要一個控制項。 **Control** 元素可以是 **Button** 或 **Menu**。 使用 **Menu** 指定按鈕控制項的下拉式清單。 目前僅支援按鈕和功能表。請參閱[按鈕控制項](#按鈕控制項)和[功能表控制項](#功能表控制項)節以取得詳細資訊。<br/>**附註**  若要使疑難排解更加容易，我們建議一次加入一個 **Control** 元素和相關的 **Resources** 子元素。

## Outlook 增益集命令的擴充點

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (只能在 [DesktopFormFactor](./formfactor.md) 中使用。)

### CustomPane

CustomPane 擴充點定義了符合指定的規則時所啟動的增益集。 它僅適用於讀取表單，並顯示在水平窗格中。 

**子元素**

|  元素 |  必要  |  說明  |
|:-----|:-----|:-----|
|  **RequestedHeight** | 否 |  要求的高度 (單位為像素)，適用於桌上型電腦上執行的顯示窗格。 這可以從 32 到 450 個像素。  |
|  **SourceLocation**  | 是 |  增益集的原始程式碼檔的 URL。 這是指 [Resources](./resources.md) 元素中的 **Url** 元素。  |
|  **規則**  | 是 |  增益集啟動時指定的規則或規則集合。 如需詳細資訊，請參閱 [Outlook 增益集的啟用規則](../../outlook/manifests/activation-rules.md)。 |
|  **DisableEntityHighlighting**  | 否 |  指定是否要關閉實體醒目提示功能。 |


#### CustomPane 範例
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```

### MessageReadCommandSurface
這個擴充點會將按鈕放在郵件讀取檢視的命令介面中。 在 Outlook 桌面中，這將會出現在功能區。

**子元素**

|  元素 |  說明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

#### OfficeTab 範例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab 範例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### MessageComposeCommandSurface
這個擴充點會使用郵件撰寫表單，將增益集的按鈕放在功能區上。 

**子元素**

|  元素 |  說明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

#### OfficeTab 範例
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab 範例

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### AppointmentOrganizerCommandSurface

這個擴充點會將向會議召集人顯示的表單按鈕放在功能區中。 

**子元素**

|  元素 |  說明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

#### OfficeTab 範例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab 範例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

這個擴充點會將向會議出席者顯示的表單按鈕放在功能區中。 

**子元素**

|  元素 |  說明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

#### OfficeTab 範例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### CustomTab 範例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### 模組

這個擴充點會將按鈕放在模組擴充的功能區上。 

**子元素**

|  元素 |  說明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  將命令新增至預設的功能區索引標籤中。  |
|  [CustomTab](./customtab.md) |  將命令新增至自訂的功能區索引標籤中。  |

