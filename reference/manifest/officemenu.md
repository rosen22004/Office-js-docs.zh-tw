# <a name="officemenu-element"></a>OfficeMenu 元素
定義要新增至 Office 內容功能表的控制項的集合。適用於 Word、Excel、PowerPoint 及 OneNote 增益集。

## <a name="attributes"></a>屬性

| 屬性            | 必要 | 描述                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | 是      | 正在定義的 OfficeMenu 類型。|

## <a name="child-elements"></a>子元素
|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  [Control](#control)    | 是 |  一或多個控制項物件的集合  |

## <a name="xsi:type"></a>xsi:type
指定要為其新增此 Office 增益集的 Office 用戶端應用程式的內建功能表。

- `ContextMenuText` - 選取文字，然後使用者在選取的文字上開啟內容功能表 (按一下滑鼠右鍵) 時，會顯示內容功能表上的項目。適用於 Word、Excel、PowerPoint 及 OneNote。
- `ContextMenuCell` - 當使用者在試算表的儲存格上開啟內容功能表 (按一下滑鼠右鍵) 時，會在內容功能表上顯示項目。適用於 Excel。 

## <a name="control"></a>控制項

在一或多個 [menu](./menu.md#menu-control) 控制項上需要**OfficeMenu** 


## <a name="example"></a>範例

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
