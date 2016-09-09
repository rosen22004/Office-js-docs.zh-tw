# Hosts 元素

指定啟動 Office 增益集的 Office 用戶端應用程式。 包含 **Host** 元素及其設定的集合。 

包含在 [VersionOverrides](./versionoverrides.md) 節點中時，此元素會覆寫資訊清單的父部分中的 **Hosts** 元素。 

## 子元素

|  元素 |  必要  |  說明  |
|:-----|:-----|:-----|
|  [主應用程式](#主應用程式)    |  是   |  說明主應用程式及其設定。 |

> ** 附註：** Outlook 需要 `Hosts` 包含 `MailHost` 的 `Host` 定義。

---- 

## Host 元素
指定應該啟動增益集的個別 Office 應用程式類型，例如「文件」、「活頁簿」、「簡報」、「專案」、「信箱」和「notebook」。

### 屬性

|  屬性  |  必要  |  說明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 說明套用這些設定的 Office 主應用程式。|

### 子元素

|  元素 |  必要  |  說明  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  是   |  定義受影響的表單係數。 |


### xsi:type
控制也套用包含的設定的 Office 主應用程式 (Word、Excel、PowerPoint、Outlook、OneNote)。 此值必須是下列任一項：

- `MailHost` (Outlook)    


## 主應用程式範例 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
