# <a name="mobileformfactor-element"></a>MobileFormFactor 元素

為行動裝置外觀尺寸的增益集指定設定。除了 **Resources** 節點外，它包含行動裝置外觀尺寸的所有增益集資訊。

> **附註：**目前只有 iOS 版 Outlook 支援 **MobileFormFactor** 元素。

每個 **MobileFormFactor** 定義都包含 **FunctionFile** 元素與一或多個 **ExtensionPoint** 元素。如需詳細資訊，請參閱 [FunctionFile 元素](./extensionpoint.md)和 [ExtensionPoint 元素](./functionfile.md)。

**MobileFormFactor** 元素定義於 VersionOverrides schema 1.1 中。抑制 [VersionOverrides](./versionoverrides.md) 元素必須具備 `VersionOverridesV1_1` 的 `xsi:type` 屬性值。

## <a name="child-elements"></a>子元素

| 元素                               | 必要 | 描述  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | 是      | 定義增益集公開功能的位置。 |
| [FunctionFile](./functionfile.md)     | 是      | 檔案中包含 JavaScript 函式的 URL。|

## <a name="mobileformfactor-example"></a>MobileFormFactor 範例

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
