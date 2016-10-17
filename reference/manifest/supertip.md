## <a name="supertip"></a>Supertip
定義豐富的工具提示 (包括標題和描述)。它由 [Button](./button.md) 和 [Menu](./menu-control.md) 兩個控制項所用。 

## <a name="child-elements"></a>子元素
|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  [Title](#title)        | 是 |   特別提示的文字。         |
|  [描述](#description)  | 是 |  特別提示的描述。    |

## <a name="title"></a>標題
必要。特別提示的文字。**resid** 屬性必須設定為  **Resources** 元素的 **ShortStrings** 元素中，[String](./resources.md#shortstrings) 元素的 [id](./resources.md) 屬性值。

## <a name="description"></a>描述
必要。特別提示的描述。**resid** 屬性必須設定為  **Resources** 元素的 **LongStrings** 元素中，[String](./resources.md#longstrings) 元素的 [id](./resources.md) 屬性值。

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```