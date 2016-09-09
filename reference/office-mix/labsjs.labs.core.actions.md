
# LabsJS.Labs.Core.Actions
提供 LabJS.Labs.Core.Actions JavaScript API 的概觀。

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

這些 API 代表實驗室的作業，指出實驗室的目前行為。如果您要建立新元件，或開發新驅動程式 (非 Office Mix) 的連線，則 API 非常有用。

## LabsJS.Labs.Core.Actions API 模組

Actions 模組包含下列類型︰


### 介面


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](../../reference/office-mix/labs.core.actions.iclosecomponentoptions.md)|要關閉的元件。|
|[Labs.Core.Actions.ICreateAttemptOptions](../../reference/office-mix/labs.core.actions.icreateattemptoptions.md)|嘗試的關聯元件。|
|[Labs.Core.Actions.ICreateAttemptResult](../../reference/office-mix/labs.core.actions.icreateattemptresult.md)|建立指定元件之嘗試的結果。|
|[Labs.Core.Actions.ICreateComponentOptions](../../reference/office-mix/labs.core.actions.icreatecomponentoptions.md)|建立新元件。|
|[Labs.Core.Actions.ICreateComponentResult](../../reference/office-mix/labs.core.actions.icreatecomponentresult.md)|建立新元件的 [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) 結果。|
|[Labs.Core.Actions.IGetValueResult](../../reference/office-mix/labs.core.actions.igetvalueresult.md)|取得值動作的結果。|
|[Labs.Core.Actions.ISubmitAnswerResult](../../reference/office-mix/labs.core.actions.isubmitanswerresult.md)|提交嘗試答案的結果。|
|[Labs.Core.Actions.IAttemptTimeoutOptions](../../reference/office-mix/labs.core.actions.iattempttimeoutoptions.md)|用於目前嘗試之逾時動作的選項。|
|[Labs.Core.Actions.IGetValueOptions](../../reference/office-mix/labs.core.actions.igetvalueoptions.md)|取得值作業的可用選項。|
|[Labs.Core.Actions.IResumeAttemptOptions](../../reference/office-mix/labs.core.actions.iresumeattemptoptions.md)|繼續嘗試的相關選項。|
|[Labs.Core.Actions.ISubmitAnswerOptions](../../reference/office-mix/labs.core.actions.isubmitansweroptions.md)|用於提交答案動作的選項。|

### 變數


|||
|:-----|:-----|
| `var CloseComponentAction: string`|關閉元件，並指出對它沒有進一步動作。|
| `var CreateAttemptAction: string`|用來建立新嘗試的動作。|
| `var CreateComponentAction: string`|用來建立新元件的動作。|
| `var AttemptTimeoutAction: string`|嘗試逾時動作。|
| `var GetValueAction: string`|用來擷取嘗試相關值的動作。|
| `var ResumeAttemptAction: string`|繼續嘗試動作。用來指示使用者正在指定嘗試上繼續工作。|
| `var SubmitAnswerAction: string`|用來提交指定嘗試答案的動作。|
