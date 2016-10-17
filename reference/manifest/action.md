# <a name="action-element"></a>動作元素
 指定使用者選取[按鈕](./button-control.md)或[功能表](./menu-control.md)控制項時所執行的動作。
 
## <a name="attributes"></a>屬性

|  屬性  |  必要  |  描述  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 要採取的動作類型|


## <a name="child-elements"></a>子元素

|  元素 |  描述  |
|:-----|:-----|
|  [FunctionName](#functionname) |    指定要執行的函式名稱。 |
|  [SourceLocation](#sourcelocation) |    指定此動作的來源檔案位置。 |
  

## <a name="xsi:type"></a>xsi:type
這個屬性會指定當使用者選取按鈕時執行的動作種類。它可以是下列其中一項：
- ExecuteFunction
- ShowTaskpane

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
