
# <a name="labsjs.labs.components"></a>LabsJS.Labs.Components
提供 Labs.JS Labs.Components JavaScript API 的高階概觀。

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

Labs.Components 模組中的 API 代表實驗室開發目前可用的四個預設元件 (活動、選擇、輸入和動態元件)。

## <a name="labs.components-module"></a>Labs.Components 模組

以下是 Labs.Components 類型︰


### <a name="classes"></a>類別


|||
|:-----|:-----|
|[Labs.Components.ComponentAttempt](../../reference/office-mix/labs.components.componentattempt.md)|元件上嘗試的基底類別。|
|[Labs.Components.ActivityComponentAttempt](../../reference/office-mix/labs.components.activitycomponentattempt.md)|表示完成活動元件時的嘗試。|
|[Labs.Components.ActivityComponentInstance](../../reference/office-mix/labs.components.activitycomponentinstance.md)|表示活動元件的目前執行個體。|
|[Labs.Components.ChoiceComponentAnswer](../../reference/office-mix/labs.components.choicecomponentanswer.md)|選擇元件中所呈現的問題答案。|
|[Labs.Components.ChoiceComponentAttempt](../../reference/office-mix/labs.components.choicecomponentattempt.md)|表示選擇元件上的嘗試。|
|[Labs.Components.ChoiceComponentInstance](../../reference/office-mix/labs.components.choicecomponentinstance.md)|表示選擇元件上的執行個體。|
|[Labs.Components.ChoiceComponentResult](../../reference/office-mix/labs.components.choicecomponentresult.md)|選擇元件提交的結果。|
|[Labs.Components.ChoiceComponentSubmission](../../reference/office-mix/labs.components.choicecomponentsubmission.md)|表示與選擇元件相關聯的提交。|
|[Labs.Components.DynamicComponentInstance](../../reference/office-mix/labs.components.dynamiccomponentinstance.md)|代表動態元件的執行個體。|
|[Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)|表示輸入元件問題的答案。|
|[Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md)|表示嘗試與輸入元件互動。|
|[Labs.Components.InputComponentInstance](../../reference/office-mix/labs.components.inputcomponentinstance.md)|代表輸入元件的執行個體。|
|[Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)|輸入元件提交的結果。|
|[Labs.Components.InputComponentSubmission](../../reference/office-mix/labs.components.inputcomponentsubmission.md)|表示提交到輸入元件。|

### <a name="interfaces"></a>介面


|||
|:-----|:-----|
|[Labs.Components.IActivityComponent](../../reference/office-mix/labs.components.iactivitycomponent.md)|表示活動元件。展開 [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)。|
|[Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md)|表示活動元件的特定執行個體。展開 [Labs.Core.IComponentInstance](../../reference/office-mix/labs.core.icomponentinstance.md)。|
|[Labs.Components.IChoice](../../reference/office-mix/labs.components.ichoice.md)|指定問題的可用選擇。|
|[Labs.Components.IChoiceComponent](../../reference/office-mix/labs.components.ichoicecomponent.md)|可與選擇元件互動。|
|[Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md)|選擇元件的執行個體。|
|[Labs.Components.IDynamicComponent](../../reference/office-mix/labs.components.idynamiccomponent.md)|可與動態元件互動。|
|[Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md)|動態元件的執行個體。|
|[Labs.Components.IHint](../../reference/office-mix/labs.components.ihint.md)|實驗室問題的提示。|
|[Labs.Components.IInputComponent](../../reference/office-mix/labs.components.iinputcomponent.md)|啟用與輸入元件的互動。|
|[Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md)|輸入元件的執行個體。|
