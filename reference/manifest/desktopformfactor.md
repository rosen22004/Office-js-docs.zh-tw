# <a name="desktopformfactor-element"></a>DesktopFormFactor 元素

為桌面外觀尺寸的增益集指定設定。桌面外觀尺寸包括 Office for Windows、Office for Mac 和 Office Online。除了 **Resources** 節點外，它包含桌面外觀尺寸的所有增益集資訊。

各個 DesktopFormFactor 包含 **FunctionFile** 元素與一或多個 **ExtensionPoint** 元素。如需詳細資訊，請參閱 [FunctionFile 元素](./functionfile.md)和 [ExtensionPoint 元素](./extensionpoint.md)。 

## <a name="child-elements"></a>子元素

| 元素                               | 必要 | 描述  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | 是      | 定義增益集公開功能的位置。 |
| [FunctionFile](./functionfile.md)     | 是      | 檔案中包含 JavaScript 函式的 URL。|
| [GetStarted](./getstarted.md)         | 不可以       | 定義在 Word、Excel 或 PowerPoint 主應用程式中安裝增益集時，顯示的圖說文字。 |

## <a name="desktopformfactor-example"></a>DesktopFormFactor 範例

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
