
# <a name="javascript-api-for-office-reference"></a>JavaScript API for Office 參考

JavaScript API for Office 可讓您建立 Web 應用程式，與 Office 主應用程式中的物件模型互動。您的應用程式會參考 office.js 程式庫，也就是指令碼載入器。office.js 程式庫會載入執行增益集的 Office 應用程式所適用的物件模型。您可以使用下列 JavaScript 物件模型︰


1. 一般 API - Office 2013 導入的 API。這是為**所有 Office 主應用程式**載入，並連接您的增益集應用程式與 Office 用戶端應用程式。物件模型包含 Office 用戶端的特定 API，且適用於多個 Office 用戶端主應用程式的 API。此內容的全部都在**共用 API** 底下。**Outlook** 也會使用一般 API 語法。程式碼中別名 [Office](../reference/shared/office.md) 下的所有項目，包含您可用來撰寫指令碼的物件，以便與 Office 文件、工作表、簡報、郵件項目，以及 Office 增益集的專案中內容互動。如果增益集將選定 Office 2013 及更新版本，您必須使用這些常見的 API。這個物件模型會使用回呼。

1. **Office 2016** 導入的主機特定 API。這個物件模型提供主機特定的強型別物件，其對應於您使用 Office 用戶端時會看到的熟悉物件，並且代表 Office JavaScript API 的未來。主機特定 API 目前包含 [Word JavaScript API](../reference/word/word-add-ins-reference-overview.md) 和 [Excel JavaScript API](../reference/excel/application.md)。

從 TOC 上方的下拉式清單選取 Office 用戶端，根據您的目標主應用程式篩選內容。

## <a name="supported-host-applications"></a>支援的主應用程式
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word

深入了解[支援的主機及其他需求](../docs/overview/requirements-for-running-office-add-ins.md)。

## <a name="open-api-specifications"></a>開放式 API 規格

我們設計和開發新的 Office 增益集 API 時，我們會將其放在[開放式 API 規格](openspec.md)頁面中，可供您提出意見反應。了解即將推出的新功能，並對我們的設計規格提出意見反應。

