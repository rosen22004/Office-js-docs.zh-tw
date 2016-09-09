
# 設定及編輯 Office Mix LabsJS 實驗室



Office Mix 提供 office.js 方法來取得及設定實驗室組態。組態表示 Office Mix 您正在建立的實驗室類型，以及實驗室會傳回的資料類型。此資訊是用來收集和視覺化分析。

## 取得實驗室編輯器

實驗室編輯器 [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md) 物件可讓您編輯實驗室，及取得和設定您的實驗室組態。當您完成實驗室編輯時，您必須呼叫 **Done** 方法。不過，除非您正在嘗試取得或執行您正在編輯的實驗室，否則不需要呼叫 **Done** 方法。請注意，一次只能開啟實驗室的一個執行個體。

下列程式碼示範如何取得實驗室編輯器。




```js
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

使用 **Labs.LabEditor** 上的 **getConfiguration** 和 [setConfiguration](../../../reference/office-mix/labs.labeditor.md) 方法來儲存指定實驗室的組態。組態 ([Labs.Core.IConfiguration](../../../reference/office-mix/labs.core.iconfiguration.md)) 指示 Office Mix 實驗室將收集和處理哪些資料。組態包含實驗室的一般資訊，包括名稱、版本和其他組態選項。組態的最重要部分是實驗室元件的定義。

下列程式碼會示範如何設定和取得組態。若要設定組態，只要建立組態物件，然後呼叫 **setConfiguration** 方法。若要再擷取組態，請在實驗室編輯器物件上呼叫 **getConfiguration** 方法。




```js

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```js

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## 關閉編輯器

若要關閉編輯器，在完成編輯實驗室時，在編輯器上呼叫 **Done** 方法。請注意，您無法同時取得和編輯實驗室。但在呼叫 **Done** 後，您可以接著編輯或執行實驗室。


## 與實驗室互動

設定實驗室組態之後，即準備好開始與實驗室互動。在 PowerPoint 內執行實驗室時會模擬互動。不過，在 Office Mix 課程播放程式內執行實驗室時，資料是儲存在 Office Mix 資料庫並用於分析。


### 取得實驗室執行個體

您使用 [Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md) 物件 (也就是為目前使用者設定之實驗室的執行個體) 與實驗室互動。若要執行 (或「取得」) 實驗室，請呼叫 [Labs.takeLab](../../../reference/office-mix/labs.takelab.md) 函式。


```js
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

執行個體物件包含對應到組態中指定之元件的元件執行個體 ([Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md)、[Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)) 陣列。事實上，執行個體只是轉換的組態版本，用來將伺服器端 ID 附加到執行個體物件，以及向使用者隱藏適用的特定欄位 (例如，提示、解答等等)。


### 管理狀態

狀態是執行指定實驗室的使用者所相關聯的暫時儲存區。您可以使用存放區來保存實驗室的後續引動過程之間的資訊。例如，程式設計的實驗室無法儲存使用者目前正在進行的工作。

若要**設定**狀態，請使用下列程式碼。




```js
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

若要**取得**狀態，請使用下列程式碼。




```js
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## 元件執行個體和結果

以下是如何實作四種元件類型之執行個體的概觀，以及元件方法的簡短範例。 

不過，在使用元件執行個體時，您需要先熟悉兩個核心概念。第一個概念是**嘗試** 和**值**。

 **嘗試**

嘗試是使用者完成元件執行個體的嘗試。例如，若是複選問題，當使用者開始處理問題並在指派最後分數結束時，會啟動嘗試。Office Mix 分析接著會彙總使用者問題的結果。


 >**附註**：嘗試可使用於所有元件類型，但 **DynamicComponent** 類型除外。

您可以使用 **getAttempts** 方法來擷取與指定的元件執行個體相關聯的所有嘗試結果。擷取結果後，使用者可以使用 **resume** 方法來重試其中一個現有的嘗試，或使用 **createAttempt** 方法來建立新嘗試。下列範例會顯示程序。




```js
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **值**

元件執行個體包含對應到值陣列的機碼字典。您可以使用陣列來儲存提示、意見反應或任何其他一組您想要關聯至元件的值。元件執行個體使用 **getValues** 方法來提供這些值的存取。

例如，查詢提示值會導致分析為該使用者標記提示。會以每次嘗試為基礎追蹤值。

下列程式碼範例顯示如何查詢提示。




```js
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### ActivityComponentInstance


使用 **ActivityComponentInstace** 物件來追蹤使用者與活動元件的互動。這個類別會提供 **complete** 方法來表示使用者已完成與活動的互動。此方法可指出使用者已完成分派的任務、已完成讀取，或與活動關聯的任何其他端點。下列程式碼示範如何使用 **complete** 方法。


```js
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### ChoiceComponentInstance


使用 **ChoiceComponentInstance** 物件來追蹤使用者與選擇元件的互動。選擇元件是問題，可向使用者呈現他們需要從中選取的選擇清單。可能或可能沒有正確的答案。此類別提供兩種主要方法︰**getSubmissions** 和 **submit**。**getSubmissions** 方法可讓您擷取先前儲存的提交；**submit** 方法可儲存新提交。下列程式碼範例說明方法的使用。


```js
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```js
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### InputComponentInstance


使用 **InputComponentInstance** 物件來追蹤使用者與輸入元件的互動。此類別提供兩種主要方法︰**getSubmission** 和 **submit**。**getSubmissions** 方法可讓您擷取先前儲存的提交；**submit** 方法可讓您儲存新提交。下列程式碼片段說明 **getSubmissions** 方法的使用。


```js
var submissions = this._attempt.getSubmissions();
```

使用 **submit** 方法時，請注意，**InputComponentAnswer** 物件代表已提交的答案，**InputComponentResult** 物件則包含結果。傳回的值是 **InputComponentSubmission** 物件，包含答案、結果，和表示結果提交時間的時間戳記。




```js
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### DynamicComponentInstance


使用 **DynamicComponentInstance** 物件來追蹤使用者與動態元件的互動。這個類別中的主要方法是 **getComponents**、**createComponent** 和 **close**。

**GetComponents** 方法可讓您擷取先前建立的元件執行個體清單，如下列範例所示。




```js
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

**CreateComponent** 方法會建構新的元件，並傳回該元件執行個體，如下列範例所示。




```js
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

使用 **close** 方法，表示您已完成使用動態元件來建立新元件。請注意，您也可以使用 **isClosed** 布林值方法來測試是否已關閉動態元件執行個體。下列程式碼示範如何使用 **close** 方法。




```js
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## 其他資源



- [Office Mix 增益集](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [逐步解說︰建立第一個 Office Mix 實驗室](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
