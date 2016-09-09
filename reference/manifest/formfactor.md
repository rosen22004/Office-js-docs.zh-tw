# FormFactor 元素

為指定外觀尺寸的增益集指定設定。 例如，使用型別 `MailHost` 與 `DesktopFormFactor` 定義 `Host` 會套用到桌面的 Outlook，但_不會_套用到 Outlook Web App 或 Outlook.com。 除了 **Resources** 節點外，它包含該外觀尺寸的所有增益集資訊。

外觀尺寸包含 **FunctionFile** 元素與一或多個 **ExtensionPoint** 元素。 如需詳細資訊，請參閱下列 [FunctionFile element](./functionfile.md)和 [ExtensionPoint element](./extensionpoint.md)一節。 

支援 FormFactors 如下：

- `DesktopFormFactor` (Office for Windows 或 Mac 用戶端)

## 子元素

| 元素                               | 必要 | 說明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | 是      | 定義增益集公開功能的位置。 |
| [FunctionFile](./functionfile.md)     | 是      | 檔案中包含 JavaScript 函式的 URL。|
| [GetStarted](./getstarted.md)         | 否       | 定義在 Word、Excel 或 PowerPoint 主應用程式中安裝增益集時，顯示的圖說文字。 |

## FormFactor 範例

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
