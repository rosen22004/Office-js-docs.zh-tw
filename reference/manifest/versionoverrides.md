# <a name="versionoverrides-element"></a>VersionOverrides 元素

根元素，其包含增益集實作的增益集命令的資訊。**VersionOverrides** 是資訊清單中 [OfficeApp](./officeapp.md) 元素的子元素。該元素是在資訊清單結構描述 1.1 版及更新版本支援，但在 VersionOverrides v1.0 結構描述中定義。 

## <a name="attributes"></a>屬性

|  屬性  |  必要  |  描述  |
|:-----|:-----|:-----|
|  **xmlns**       |  是  |  結構描述位置，必須是 `http://schemas.microsoft.com/office/mailappversionoverrides`。|
|  **xsi:type**  |  是  | 結構描述版本。目前唯一有效的值是 `VersionOverridesV1_0`。 |


## <a name="child-elements"></a>子元素

|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  **描述**    |  不可以   |  描述增益集。這會覆寫資訊清單中任何父部份的 `Description` 元素。描述文字內含於 [Resources](./resources.md) 元素的 **LongString** 元素的子元素中。**Description** 元素的 `resid` 屬性會設定為包含文字的 `String` 元素的 `id` 屬性值。|
|  **需求**  |  不可以   |  指定增益集需要的 Office.js 的最小需求集和版本。這會覆寫資訊清單中父部份的 `Requirements` 元素。| 
|  [主機](./hosts.md)                |  是  |  指定 Office 主機的集合。此子 Hosts 元素會覆寫資訊清單的父部分中的 Hosts 元素。  |
|  [資源](./resources.md)    |  是  | 定義其他資訊清單元素參考的資源 (字串、URL 和影像)。|



### <a name="versionoverrides-example"></a>VersionOverrides 範例
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information on resources -->
   </Resources>
</VersionOverrides>
...
</OfficeApp>
```
