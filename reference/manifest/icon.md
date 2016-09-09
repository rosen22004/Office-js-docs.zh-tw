# 圖示元素
定義[按鈕](./button.md)和[功能表](./menu-control.md)控制項的**影像**元素。

## 子元素
|  元素 |  必要  |  說明  |
|:-----|:-----|:-----|
|  [影像](#影像)        | 是 |   要使用的影像的 resid         |

## 影像
按鈕的影像。**resid** 屬性必須設定為 **Resources** 的 **Images** 元素中，**Image** 元素的 [id](./resources.md) 屬性值。**size** 屬性指出影像的大小，單位為像素。需要三個影像大小 (16、32 與 80 像素)，但支援其他五個大小 (20、24、40、48 及 64 像素)。|


```xml
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
```  