
# <a name="labs.components.dynamiccomponentinstance"></a>Labs.Components.DynamicComponentInstance

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

代表動態元件的執行個體。

```
class DynamicComponentInstance extends Labs.ComponentInstanceBase
```


## <a name="properties"></a>屬性


|屬性	|描述|
|:-----|:-----|
| `public var component: Components.IDynamicComponentInstance`|元件執行個體定義。|

## <a name="methods"></a>方法




### <a name="constructor"></a>建構函式

 `function constructor(component: Components.IDynamicComponentInstance)`

使用 [Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md) 定義來建立新的動態元件執行個體。


### <a name="getcomponents"></a>getComponents

 `public function getComponents(callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase[]>): void`

擷取這個動態元件建立的所有元件。

 **參數**


|參數|描述|
|:-----|:-----|
| _callback_|擷取所有元件後所引發的回呼函式。|

### <a name="createcomponent"></a>createComponent

 `public function createComponent(component: Labs.Core.IComponent, callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase>): void`

將動態元件作為基底元件來建立新元件。

 **參數**


|參數|描述|
|:-----|:-----|
| _component_|建立執行個體的來源元件 ([Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md))。|
| _callback_|建立元件後所引發的回呼函式。|

### <a name="close"></a>close

 `public function close(callback: Labs.Core.ILabCallback<void>): void`

表示這個元件執行個體沒有相關聯的額外提交。

 **參數**


|參數|描述|
|:-----|:-----|
| _callback_|執行個體關閉後所引發的回呼函式。|

### <a name="isclosed"></a>isClosed

 `public function isClosed(callback: Labs.Core.ILabCallback<boolean>): void`

傳回是否要關閉動態元件。如果關閉則傳回 **true**。

