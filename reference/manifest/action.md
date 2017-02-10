# <a name="action-element"></a>動作元素
指定使用者選取[按鈕](./control.md#button-control)或[功能表](./control.md#menu-dropdown-button-controls)控制項時所執行的動作。
 
## <a name="attributes"></a>屬性

|  屬性  |  必要  |  描述  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 要採取的動作類型|


## <a name="child-elements"></a>子元素

|  元素 |  描述  |
|:-----|:-----|
|  [FunctionName](#functionname) |    指定要執行的函式名稱。 |
|  [SourceLocation](#sourcelocation) |    指定此動作的來源檔案位置。 |
|  [TaskpaneId](#taskpaneid) | 指定工作窗格容器的識別碼。|
|  [SupportsPinning](#supportspinning) | 指定可支援釘選的工作窗格，其能在使用者變更選取項目時，仍維持工作窗格的開啟狀態。|
  

## <a name="xsitype"></a>xsi:type
這個屬性會指定當使用者選取按鈕時執行的動作種類。它可以是下列其中一項：

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

當 **xsi:type** 為 "ExecuteFunction" 的必要元素。指定要執行的函式名稱。函式內含於 [FunctionFile](./functionfile.md) 元素中指定的檔案。

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation
當 **xsi:type** 為 "ShowTaskpane" 的必要元素。指定此動作的來源檔案位置。**resid** 屬性必須設定為 **Resources** 的 **Urls** 元素中，[Url](./resources.md#urls) 元素的 [id](./resources.md) 屬性值。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId
當 **xsi:type** 為「ShowTaskpane」的選擇性元素。指定工作窗格容器的識別碼。如果您想要每項皆有獨立的窗格，在您有多個「ShowTaskpane」的動作時，使用不同的 **TaskpaneId**。使用相同 **TaskpaneId** 為不同的動作，共用相同的窗格。當使用者選擇共用相同的命令 **TaskpaneId** 時，窗格容器會保持開啟，但對應的動作「SourceLocation」將會取代窗格的內容。 

>**附註︰**Outlook 中已不支援此項目。

下列範例顯示兩個「動作」共用相同的 **TaskpaneId**。 


```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

## <a name="supportspinning"></a>SupportsPinning

當 **xsi:type** 為「ShowTaskpane」的選擇性元素。包含 [VersionOverrides](./versionoverrides.md) 的元素必須具備 `VersionOverridesV1_1` 的 `xsi:type` 屬性值。加入此項附有 `true` 值的元素以支援工作窗格釘選。使用者將能夠「釘選」工作窗格，在變更選取項目時，使其仍維持開啟狀態。如需詳細資訊，請參閱[在 Outlook 中實作可釘選的工作窗格](../../docs/outlook/manifests/pinnable-taskpane)。

> **附註**：目前僅有 Windows 版 Outlook 2016 (組建 7628.1000 或更新版本) 支援 SupportsPinning。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```