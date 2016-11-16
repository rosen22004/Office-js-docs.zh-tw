# <a name="tips-for-creating-office-addins-with-angular-2"></a>使用 Angular 2 建立 Office 增益集的祕訣 

本文提供使用 Angular 2 建立 Office 增益集做為單一頁面應用程式的指引。

>**附註：**根據您使用 Angular 2 來建立 Office 增益集的經驗，您是否可提供一些建議？您可以在 [GitHub](https://github.com/OfficeDev/office-js-docs) 中參與本文討論，或在儲存機制中提交[問題](https://github.com/OfficeDev/office-js-docs/issues)以提供意見反應。 

如需使用 Angular 2 架構建置的 Office 增益集範例，請參閱[在 Angular 2 上建置的 Word 樣式檢查增益集](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>啟動載入必須在 Office.initialize 內

在任何呼叫 Office、Word 或 Excel JavaScript API 的頁面上，您的程式碼必須先指派方法給 `Office.initialize` 屬性。(如果您沒有初始化程式碼，方法主體可能只是空的 "`{}`" 符號，但您必須未定義 `Office.initialize` 屬性。如需詳細資訊，請參閱[初始化增益集](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in))。Office 會在初始化 Office JavaScript 程式庫之後，立即呼叫這個方法。

**Angular 啟動載入程式碼必須在您指派給 `Office.initialize`** 的方法內呼叫，以確保先初始化 Office JavaScript 程式庫。以下是顯示如何執行這項操作的簡單範例。此程式碼應位於專案的 main.ts 檔案中。

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
    import { AppModule } from './app.module';
    Office.initialize = function () {
        const platform = platformBrowserDynamic();
        platform.bootstrapModule(AppModule);
  };
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>在 Angular 應用程式中使用雜湊位置策略

如果未指定雜湊位置策略，則可能無法在應用程式中的路由之間巡覽。您可以利用下列兩個方法的其中之一來完成這項作業。首先，您可以指定應用程式模組中位置策略的提供者，如下列範例所示。它會進入 app.module.ts 檔案中。

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity
    @NgModule({
        providers: [
            {provide: LocationStrategy, useClass: HashLocationStrategy},
            // Other providers suppressed
        ],
        // Other module properties suppressed
  })
  export class AppModule {}
``` 

如果您在不同的路由模組中定義路由，則沒有替代方法可指定雜湊位置策略。在路由模組的 .ts 檔案中，將組態物件傳遞至可指定此策略的 `forRoot` 函式。下列程式碼是一個範例。 

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity
    const routes: Routes = // route definitions go here
    @NgModule({
      imports: [ RouterModule.forRoot(routes, {useHash: true}) ],
      exports: [ RouterModule ]
    })
    export class AppRoutingModule {}
```   


## <a name="consider-wrapping-fabric-components-with-angular-2-components"></a>請考慮使用 Angular 2 元件包裝 Fabric 元件

我們建議在增益集中使用 [Office UI Fabric](http://dev.office.com/fabric#/fabric-js) 樣式。Fabric 包括數個版本所隨附的元件，包括[以 TypeScript 為基礎](https://github.com/OfficeDev/office-ui-fabric-js)的版本。請考慮在 Angular 2 元件中包裝 Fabric 元件，以便在增益集中使用這些元件。如需示範如何執行這項操作的範例，請參閱[在 Angular 2 上建置的 Word 樣式檢查增益集](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。請注意，舉例來說，[fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) 中的 Angular 元件定義方式可匯入 Fabric 檔案 TextField.ts (Fabric 元件定義所在)。 


## <a name="using-the-office-dialog-api-with-angular"></a>搭配使用 Office 對話方塊 API 與 Angular

Office 增益集對話方塊 API 可讓您的增益集在半強制回應對話方塊中開啟頁面，以便與主頁面 (通常位於工作窗格中) 交換資訊。 

[displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) 方法採用的參數可指定應在此對話方塊中開啟的頁面 URL。增益集可以有個別的 HTML 網頁 (與基礎頁面不同) 來傳遞至這個參數，或者您可以在 Angular 應用程式中傳遞路由的 URL。 

請務必記住，如果您傳遞路由，對話方塊會以它自己的執行內容建立新的視窗。基礎網頁與其所有的初始化和啟動載入程式碼都會在此新的內容中重新執行，而所有變數會在對話方塊中設定為其初始值。因此這項技術會在對話方塊中啟動單一頁面應用程式的第二個執行個體。在對話方塊中變更變數的程式碼不會變更相同變數的工作窗格版本。同樣地，此對話方塊有自己的工作階段存放區，其無法在工作窗格中經由程式碼存取。  


## <a name="forcing-an-update-of-the-dom"></a>強制執行 DOM 更新

在任何 Angular 2 應用程式中，不會引發偶爾更新 DOM 的通知。此架構會在 `ApplicationRef` 物件上提供 `tick()` 方法，以便強制執行更新。下列程式碼是一個範例。

```js
import { ApplicationRef } from '@angular/core';
    export class MyComponent {
        constructor(private appRef: ApplicationRef) {}
        myMethod() {
            // Code that changes the DOM is here
            appRef.tick();
        }
}
``` 

## <a name="using-observables"></a>使用 Observables

Angular 2 會使用 RxJS (Reactive Extensions for JavaScript)，而 RxJS 會引入 `Observable` 和 `Observer` 物件來實作非同步處理。本節提供如何使用 `Observables` 的簡短簡介；如需詳細資訊，請參閱正式的 [RxJS](http://reactivex.io/rxjs/) 文件。

`Observable` 在某些方面很像 `Promise` 物件 - 會立即從非同步呼叫傳回，但可能要等一段時間才會解析。不過，當 `Promise` 是單一值 (可以是陣列物件) 時，`Observable` 會是物件陣列 (可能只有單一成員)。這可讓程式碼在 `Observable` 物件上呼叫[陣列方法](http://www.w3schools.com/jsref/jsref_obj_array.asp)，例如 `concat``map` 和 `filter`。 

### <a name="pushing-instead-of-pulling"></a>推送而非提取

您的程式碼會藉由將 `Promise` 物件指派給變數來進行「提取」，但 `Observable` 物件會將其值「推送」至訂閱 `Observable` 的物件。訂閱者為 `Observer` 物件。推送架構的好處是新成員可隨著時間加入至 `Observable` 陣列。加入新成員後，所有訂閱 `Observable` 的 `Observer` 物件就會收到通知。 

`Observer` 是設定為利用函式處理每個新物件 (稱為「下一頁」物件)。(它也設定為回應錯誤和完成通知。請參閱下一節中的範例)。基於這個理由，`Observable` 物件可使用於比 `Promise` 物件更廣泛的案例。例如，除了從 AJAX 呼叫傳回 `Observable` (您可以傳回 `Promise` 的方法) 以外，還可以從事件處理常式傳回 `Observable`，例如「變更」文字方塊的事件處理常式。每次使用者在方塊中輸入文字時，所有訂閱的 `Observer` 物件都會使用最新的文字及/或應用程式的目前狀態立即反應做為輸入。 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a>等到所有非同步呼叫完成為止

當您想要確定只會在一組 `Promise` 物件的每個成員都已解析時才執行回呼，請使用 `Promise.all()` 方法。

```js
myPromise.all([x, y, z]).then(// TODO: Callback logic goes here.)
``` 

若要使用 `Observable` 物件執行相同的作業，請使用 [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) 方法。  

```js
var source = Rx.Observable.forkJoin([x, y, z]);

var subscription = source.subscribe(
  function (x) {
    // TODO: Callback logic goes here
  },
  function (err) {
    console.log('Error: ' + err);
  },
  function () {
    console.log('Completed');
  });
``` 

