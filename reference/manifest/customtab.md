# CustomTab 元素
在功能區中，您可以指定其增益集命令的索引標籤和群組。 這可位於預設索引標籤 (不論是 [家用]****、[訊息]**** 或 [會議]****)，或增益集所定義的自訂索引標籤。

在 [自訂] 索引標籤上，增益集可以建立最多 10 個群組。每個群組僅限於 6 個控制項，無論其出現在哪一個索引標籤。增益集僅限於一個自訂索引標籤。

**id** 屬性在資訊清單內必須是唯一的。

## 子元素
|  元素 |  必要  |  說明  |
|:-----|:-----|:-----|
|  [群組](./group.md)      | 是 |  定義命令群組。  |
|  [標籤](#標籤)      | 是 |  CustomTab 或群組的標籤。  |
|  [控制項](#控制項)    | 是 |  一或多個控制項物件的集合  |

## 群組
必要。 請參閱[群組元素](./group.md)。

## 標籤 (索引標籤)
必要。自訂索引標籤的標籤。**resid** 屬性必須設定為  **Resources** 元素的 **ShortStrings** 元素中，[String](./resources.md#shortstrings) 元素的 [id](./resources.md) 屬性值。


##  CustomTab 範例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```