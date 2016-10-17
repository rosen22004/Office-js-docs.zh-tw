
# <a name="guidelines-for-creating-labs-for-mix-using-labsjs"></a>使用 LabsJS 建立 Mix 實驗室的指導方針



LabsJS 程式庫 (labs.js) 支援撰寫特定的 Office 增益集 (稱為實驗室)，它能整合 Office Mix。Office Mix 接下來會使用 Microsoft PowerPoint 呈現實驗室。雖然我們將這些元件稱為「實驗室」，但請注意，我們正在建立特殊的 Office 增益集，即 Office Mix 增益集。

LabsJS 內容藉由提供指南和範例來協助您實作 labs.js JavaScript API。此文件庫建置在 [JavaScript API for Office](../../../reference/javascript-api-for-office.md) (Office.js) 上方，並提供為 Office Mix 中內嵌的增益集所最佳化的抽象層。


## <a name="general-guidelines"></a>一般指導方針


以下是部分一般指導方針，可協助您使用 LabJS API 撰寫增益集。


### <a name="scripts"></a>Scripts

由於 labs.js 程式庫是 office.js 上的一個抽象層，因此會相依於 office.js，所以 office.js 和 labs.js 程式庫檔案都必須包含在開發專案中。 

您可以在 `<script src="https://sforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>` 參考 office.js 程式庫。

Labs.js 程式庫隨附 LabsJS SDK。或者，您可以在 CDN 上參考 labs.js 程式庫，網址是  <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>。請注意，實驗室的實際執行版本必須參考 CDN 上儲存的版本。


 >**附註**：除了 JavaScript 檔案 (labs-1.0.4.js) 外，我們還提供實驗室 API 的 TypeScript 定義檔 (labs-1.0.4.d.ts)。定義檔是根據 TypeScript 0.9.1.1 版所建立。


### <a name="callbacks-and-error-handling"></a>回呼和錯誤處理

Labs.js API 中的某些方法會以非同步方式操作。對於這些作業，API 會採用標準回呼介面，**ILabCallback**。 


```js
function(err, result) {
}
```

回呼方法採用兩個參數，_err_ 和 _result_。除非發生錯誤，否則 [_err]__ 欄位會保留 **null**。[結果]__ 欄位會傳回運算的結果。

即使結果立即可用，永遠不會立即引發回呼作業。相反地，它就會引發 JavaScript 事件迴圈的個別執行 (藉由 **setTimeout** 呼叫)。採用此回呼定義，您可以輕易地整合 labs.js 與您選擇的承諾 API。例如，您可以使用簡單的轉譯方法來替代這些回呼的 jQuery 承諾，如下列範例所示。




```js
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### <a name="lab-host-and-defaultlabhost"></a>實驗室主機和 DefaultLabHost

實驗室主機 ( **ILabHost**) 是支援實驗室開發的基礎驅動程式。根據預設，這是設為整合 office.js 的主機。

針對測試用途，及在 labhost.html 中執行您的實驗室，您必須切換到在模擬環境中運作的主機。下列程式碼範例顯示如何使用查詢參數執行此工作。或者，您可以變更 **DefaultHostBuilder** 以完全整合實驗室增益集與不同的平台。




```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### <a name="initialization"></a>初始化

初始化會建立實驗室及其主機之間的通訊路徑。藉由呼叫下列來初始化實驗室。


```js
Labs.connect((err, connectionResponse) => {});
```

初始化之後，您可以呼叫 labs.js API 的其他方法。_ConnectionResponse_ 參數會包含主機和使用者的相關資訊，和其他連線的相關資訊。如需傳回值的詳細資訊，請參閱 [Labs.Core.IConnectionResponse](../../../reference/office-mix/labs.core.iconnectionresponse.md)。


### <a name="time-format"></a>時間格式

Labs.js 會儲存數字，作為從 1970 年 1 月 1 日 UTC 後經過的毫秒。這符合 JavaScript [Date 物件](http://msdn.microsoft.com/en-us/library/ie/cd9w2te4%28v=vs.94%29.aspx)的日期格式，


### <a name="timeline"></a>時間表

實驗室也可與課程播放程式時間表互動。時間表可讓實驗室指示課程播放程式進入下一張投影片。可藉由呼叫 **Labs.getTimeline** 方法來擷取時間表物件。


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="handling-events"></a>處理事件


LabsJS 事件 API 會追蹤實驗室特定事件並讓您加入事件處理常式，讓您能夠回應或對事件採取行動。事件方法有三種，位於  **EventTypes** 物件︰**ModeChanged**、**Activate** 和 **Deactivate**。 


### <a name="mode-change"></a>模式變更

當指定的實驗室從編輯模式變更為檢視模式，就會引發 **ModeChanged** 事件。在 PowerPoint 編輯模式中檢視實驗室時，會顯示編輯模式。當 PowerPoint 呈現投影片放映，或當 Office Mix 課程播放程式正在顯示實驗室時，可看到檢視模式。取得實驗室時，檢視模式應永遠顯示使用者所看見的內容。編輯模式可讓使用者設定實驗室。

傳遞至回呼的 **ModeChangedEventData** 物件中資料包含目前模式的相關資訊。下列程式碼示範如何使用 **ModeChanged** 事件。




```js
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### <a name="activate"></a>啟動

當實驗室所在的 PowerPoint 投影片在課程播放程式中啟動時，就會引發**啟動**事件。


```js
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### <a name="deactivate"></a>Deactivate

當實驗室所在的 PowerPoint 投影片不再是作用中投影片時，就會引發**停用**事件。


```js
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### <a name="timeline"></a>時間表

實驗室也可與課程播放程式時間表互動。時間表可讓實驗室指示課程播放程式進入下一張投影片。可藉由呼叫 **Labs.getTimeline** 方法來擷取時間表物件。


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="additional-resources"></a>其他資源



- [Office Mix 增益集](../../powerpoint/office-mix/office-mix-add-ins.md)
    
