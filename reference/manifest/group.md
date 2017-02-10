# <a name="group-element"></a>Group 元素
定義索引標籤中的一群 UI 控制項。在 [自訂] 索引標籤上，增益集可以建立最多 10 個群組。每個群組僅限於 6 個控制項，無論其出現在哪一個索引標籤。增益集僅限於一個自訂索引標籤。

## <a name="attributes"></a>屬性

|  屬性  |  必要  |  描述  |
|:-----|:-----|:-----|
|  [id](#xsitype)  |  是  | 群組的唯一 ID。|

## <a name="child-elements"></a>子元素
|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  [Label](#label)      | 是 |  CustomTab 或群組的標籤。  |
|  [Control](#control)    | 是 |  一或多個控制項物件的集合。  |

## <a name="id-attribute"></a>id 屬性
必要。群組的唯一識別項。它是最多為 125 個字元的字串。這在資訊清單內必須是唯一的，否則群組將無法呈現。

## <a name="label"></a>標籤 
必要。群組的標籤。**resid** 屬性必須設定為  **Resources** 元素的 **ShortStrings** 元素中，[String](./resources.md#shortstrings) 元素的 [id](./resources.md) 屬性值。

## <a name="control"></a>控制項
一個群組至少需要一個控制項。

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```