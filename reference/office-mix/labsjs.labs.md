
# <a name="labsjs.labs"></a>LabsJS.Labs

 _**適用於︰**Office 的應用程式 | Office 增益集 | Office Mix | PowerPoint_

LabsJS.Labs 模組包含可用來建立 Office 增益集 (實驗室) 的機碼 JavaScript API 集。API 會提供實驗室開發的進入點。

## <a name="labsjs.labs-api-module"></a>LabsJS.Labs API 模組

Labs 模組包含下列類型︰


### <a name="variables"></a>變數


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md)|使用這個物件來建構預設的 [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) 執行個體。|

### <a name="functions"></a>函數


|||
|:-----|:-----|
|[Labs.Connect](../../reference/office-mix/labs.connect.md)|初始化與主機的連線。|
|[Labs.connect (overload)](../../reference/office-mix/labs.connect-overload.md)|初始化與主機的連線並提供輸入參數。|
|[Labs.isConnected](../../reference/office-mix/labs.isconnected.md)|初始化與主機的連線。|
|[Labs.getConnectionInfo](../../reference/office-mix/labs.getconnectioninfo.md)|擷取與指定的連線相關聯的組態資訊。|
|[Labs.disconnect](../../reference/office-mix/labs.disconnect.md)|從主機中斷連接實驗室，並提供實驗室完成狀態。|
|[Labs.editLab](../../reference/office-mix/labs.editlab.md)|開啟指定的實驗室進行編輯。您可以在編輯模式中指定實驗室的組態資料。不過，您無法編輯正在取得的實驗室 (也就是實驗室正在執行)。|
|[Labs.takeLab](../../reference/office-mix/labs.takelab.md)|執行指定的實驗室，並可將實驗室結果傳送至伺服器。請注意，您無法執行正在編輯的實驗室。|
|[Labs.on](../../reference/office-mix/labs.on.md)|新增指定事件的新處理常式。|
|[Labs.off](../../reference/office-mix/labs.off.md)|移除指定事件的事件處理常式。|
|[Labs.getTimeline](../../reference/office-mix/labs.gettimeline.md)|擷取您可以用來控制主機播放程式控制項的 [Labs.Timeline](../../reference/office-mix/labs.timeline.md) 物件執行個體。|
|[Labs.registerDeserializer](../../reference/office-mix/labs.registerdeserializer.md)|將指定的 JSON 物件還原序列化到物件。只應由元件作者使用。|

### <a name="classes"></a>類別


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](../../reference/office-mix/labs.componentinstancebase.md)|元件執行個體初始化的基底類別。|
|[Labs.ComponentInstance](../../reference/office-mix/labs.componentinstance.md)|代表元件的執行個體，為使用者在執行階段的指定元件具現化。物件包含實驗室的特定執行之元件的轉譯檢視。|
|[Labs.Command](../../reference/office-mix/labs.command.md)|用來在用戶端與主機之間傳遞訊息的一般命令。|
|[Labs.LabEditor](../../reference/office-mix/labs.labeditor.md)|**LabEditor** 物件可讓您編輯指定的實驗室，以及取得和設定實驗室所關聯的組態資料。|
|[Labs.LabInstance](../../reference/office-mix/labs.labinstance.md)|為目前的使用者設定的實驗室執行個體。使用此物件以記錄和擷取使用者的實驗室資料。|
|[Labs.Timeline](../../reference/office-mix/labs.timeline.md)|提供 labs.js 時刻表功能的存取。|
|[Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)|保留和追蹤指定的實驗室值的容器物件。值可儲存在本機或伺服器上。|

### <a name="interfaces"></a>介面


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](../../reference/office-mix/labs.getactionscommanddata.md)|可讓您擷取 [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md) 命令的相關資料。|
|[Labs.IMessageHandler](../../reference/office-mix/labs.imessagehandler.md)|可讓您定義事件處理常式的介面。|
|[Labs.ITimelineNextMessage](../../reference/office-mix/labs.itimelinenextmessage.md)|提供與 [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx) 物件互動的方式。|
|[Labs.SendMessageCommandData](../../reference/office-mix/labs.sendmessagecommanddata.md)|
  [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx) 命令的相關資料。|
|[Labs.TakeActionCommandData](../../reference/office-mix/labs.takeactioncommanddata.md)|採取動作命令的相關資料。|

### <a name="enumerations"></a>列舉


|||
|:-----|:-----|
|[Labs.ConnectionState](../../reference/office-mix/labs.connectionstate.md)|列舉實驗室至主機的可能連線狀態。|
|[Labs.ProblemState](../../reference/office-mix/labs.problemstate.md)|指定實驗室的狀態值。|
