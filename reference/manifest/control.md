# <a name="control-element"></a>Control 元素

定義執行動作或啟動工作窗格的 JavaScript 函式。**Control** 元素可以是按鈕或功能表項目。一個 [Group](group.md) 元素中至少需包含一個 **Control**。

## <a name="attributes"></a>屬性

|  屬性  |  必要  |  描述  |
|:-----|:-----|:-----|
|**xsi:type**|是|正在定義的 Control 類型。可以是 `Button`、`Menu` 或 `MobileButton`。 |
|**id**|不可以|Control 元素的 ID。最多可有 125 個字元。|

> **附註：****xsi:type** 的 `MobileButton` 值定義於 VersionOverrides 結構描述 1.1 版。它只適用於 [MobileFormFactor](./mobileformfactor.md) 元素內含的 **Control** 元素。

## <a name="button-control"></a>按鈕控制項

當使用者選取它，按鈕就會執行單一動作。它可以執行函式或顯示工作窗格。每個按鈕控制項必須有對於資訊清單唯一的 `id`。 

### <a name="child-elements"></a>子元素
|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  **標籤**     | 是 |  按鈕的文字。**resid** 屬性必須設定為 [Resources](./resources.md) 元素的 [ShortStrings](./resources.md#shortstrings) 元素中，**String** 元素的 **id** 屬性值。        |
|  **ToolTip**  |否|按鈕的工具提示。**resid** 屬性必須設為 **String** 元素的 **id** 屬性值。[Resources](resource.md) 元素的子項目是 **LongStrings** 元素，而其子項目是 **String** 元素。|     
|  [Supertip](./supertip.md)  | 是 |  按鈕的 supertip。    |
|  [Icon](./icon.md)      | 是 |  按鈕的影像。         |
|  [Action](./action.md)    | 是 |  指定要執行的動作。  |

### <a name="executefunction-button-example"></a>ExecuteFunction 按鈕範例

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a>ShowTaskpane 按鈕範例

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
## <a name="menu-dropdown-button-controls"></a>功能表 (下拉式清單按鈕) 控制項

功能表定義選項的靜態清單。每個功能表項目會執行函式或顯示工作窗格。不支援子功能表。 

當已使用 **PrimaryCommandSurface** 或 **ContextMenu** [擴充點](extensionpoint.md)後，Menu 控制項定義了︰

- 根層級功能表項目。

- 子功能表項目清單。

搭配使用 **PrimaryCommandSurface** 時，根功能表項目顯示為功能區上的按鈕。選取按鈕時，子功能表會顯示為下拉式清單。搭配使用 **ContextMenu** 時，會在快顯功能表上插入包含子功能表的功能表項目。在這兩種情況下，個別的子功能表項目可以執行 JavaScript 函式或顯示工作窗格。目前僅支援一個層級的子功能表。

下列範例示範如何定義具有兩個子功能表項目的功能表項目。第一個子功能表項目顯示工作窗格，而第二個子功能表項目執行 JavaScript 函式。

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a>子元素

|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  **標籤**     | 是 |  按鈕的文字。**resid** 屬性必須設定為  [Resources](./resources.md) 元素的 [ShortStrings](./resources.md#shortstrings) 元素中，**String** 元素的 **id** 屬性值。      |
|  **ToolTip**  |否|按鈕的工具提示。**resid** 屬性必須設為 **String** 元素的 **id** 屬性值。[Resources](resource.md) 元素的子項目是 **LongStrings** 元素，而其子項目是 **String** 元素。|     
|  [Supertip](./supertip.md)  | 是 |  這個按鈕的 supertip。    |
|  [Icon](./icon.md)      | 是 |  按鈕的影像。         |
|  [Items](#items)     | 是 |  在功能表內顯示的按鈕集合。包含每一個子功能表項目的 **Item** 元素。每個 **Item** 元素包含了[按鈕控制項](#button-control)的子元素。|

### <a name="menu-control-examples"></a>功能表控制項範例

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a>MobileButton 控制項

當使用者選取行動裝置按鈕時，它就會執行單一動作。它可以執行函式或顯示工作窗格。每個行動裝置按鈕控制項必須有對於資訊清單唯一的 `id`。

**xsi:type** 的 `MobileButton` 值定義於 VersionOverrides 結構描述 1.1 版。抑制 [VersionOverrides](./versionoverrides.md) 元素必須具備 `VersionOverridesV1_1` 的 `xsi:type` 屬性值。

> **注意**：只有 iOS 版 Outlook 支援此類型的 **Control** 元素。

### <a name="child-elements"></a>子元素
|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  **標籤**     | 是 |  按鈕的文字。**resid** 屬性必須設定為 [Resources](./resources.md) 元素的 [ShortStrings](./resources.md#shortstrings) 元素中，**String** 元素的 **id** 屬性值。        |
|  [Icon](./icon.md)      | 是 |  按鈕的影像。         |
|  [Action](./action.md)    | 是 |  指定要執行的動作。  |

### <a name="executefunction-mobile-button-example"></a>ExecuteFunction 行動裝置按鈕範例

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a>ShowTaskpane 行動裝置按鈕範例

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```