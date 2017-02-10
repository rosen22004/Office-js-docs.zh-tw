# <a name="icon-element"></a>Icon 元素
定義[按鈕](./control.md#button-control)和[功能表](./control.md#menu-dropdown-button-controls)控制項的 **Image** 元素。

## <a name="attributes"></a>屬性

|  屬性  |  必要  |  描述  |
|:-----|:-----|:-----|
|  **xsi:type**  |  否  | 正在定義的圖示類型。這只適用於行動裝置外觀尺寸的圖示。[MobileFormFactor](./mobileformfactor.md) 元素內含的 **Icon** 元素必須將此屬性設定為 `bt:MobileIconList`。 |

## <a name="child-elements"></a>子元素
|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  [Image](#image)        | 是 |   要使用的影像的 resid         |

## <a name="image"></a>影像
按鈕的影像。**resid** 屬性必須設定為 **Resources** 的 **Images** 元素中，**Image** 元素的 [id](./resources.md) 屬性值。**size** 屬性指出影像的大小，單位為像素。需要三個影像大小 (16、32 與 80 像素)，但支援其他五個大小 (20、24、40、48 及 64 像素)。|

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

### <a name="additional-requirements-for-mobile-form-factors"></a>行動裝置外觀尺寸的其他需求

當上層 **Icon** 元素 [MobileFormFactor](./mobileformfactor.md) 元素的子代時，所需的大小下限稍有不同。資訊清單至少必須提供 25、32、48 像素大小。提供的每個大小必須出現三次，其 `scale` 屬性設定為 `1`、`2` 或 `3`。

```xml
<Icon xsi:type="bt:MobileIconList">
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
```