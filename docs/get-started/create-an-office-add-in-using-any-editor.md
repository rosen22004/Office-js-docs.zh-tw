
# <a name="create-an-office-add-in-using-any-editor"></a>使用任何編輯器建立 Office 增益集

Office 增益集是您在 Office 應用程式內主控的 Web 應用程式。本文說明如何使用 Yeoman 產生器來提供專案建構和建置管理。`manifest.xml` 檔案會告訴 Office 應用程式增益集的所在位置與您想要顯示它的方式。Office 應用程式負責在 Office 內主控它。

 >**附註：**這些指示包含使用 Windows 命令提示字元的步驟，但也同樣適用於其他的 Shell 環境。 


## <a name="prerequisites-for-yeoman-generator"></a>Yeoman 產生器的必要條件

若要執行 Yeoman Office 產生器，您需要下列項目︰


- [Git](https://git-scm.com/downloads)
    
- [npm](https://www.nodejs.org/en/download)
    
- [Bower](http://bower.io/)
    
- [Yeoman Office 產生器](https://www.npmjs.com/package/generator-office)
    
- [Gulp](http://gulpjs.com/)
    
- [TSD](http://definitelytyped.org/tsd/)
    
只有 Git 和 npm 需要個別安裝。其他項目可以使用 npm 安裝。

安裝 Git 時，除了應該選擇的下列選項以外，請使用預設值︰ 


- 從 Windows 命令提示字元使用 Git
    
- 使用 Windows 的預設主控台視窗
    
使用預設值安裝 npm。然後以系統管理員身分開啟命令提示字元，並全域安裝其他軟體，可依照下列步驟執行︰




```
npm install -g bower yo generator-office gulp tsd
```


## <a name="create-the-default-files-for-your-add-in"></a>為增益集建立預設檔案

開發 Office 增益集之前，應該先為您的專案建立資料夾，並從該處執行產生器。Yeoman 產生器會在您想要建構專案的目錄中執行。 

在命令提示字元中，移至您想要建立專案的上層資料夾。然後使用下列命令來建立名為 _myHelloWorldaddin_ 的新資料夾，並將目前的目錄轉移至此處︰




```
mkdir myHelloWorldaddin
cd myHelloWorldaddin
```

使用 Yeoman 產生器來建立您所選擇的 Office 增益集︰Outlook、內容或工作窗格。在這個主題中，我們會建立工作窗格增益集。若要執行產生器，請輸入下列指示︰




```
yo office
```

產生器會提示您輸入下列︰ 


- 增益集的名稱 - 使用 _myHelloWorldaddin_
    
- 專案的根資料夾 - 使用_目前的資料夾_
    
- 增益集的類型 - 使用_工作窗格_
    
- 建立增益集所使用的技術 - _HTML、CSS &amp; JavaScript_
    
- 支援的 Office 應用程式 - 您可以選擇任何應用程式
    

**增益集的 Yeoman 產生器輸入**

![專案輸入的 yeoman 產生器提示的螢幕擷取畫面](../../images/338cf34b-fe8d-4a2f-9e38-e4bbca996139.PNG)

這會為您的增益集建立結構和基本檔案。


## <a name="hosting-your-office-add-in"></a>主控 Office 增益集

必須透過 HTTPS 提供 Office 增益集；如果是 HTTP，Office 應用程式將不會以增益集形式載入 Web 應用程式。若要在本機開發、偵錯和主控增益集，您需要一個方式在本機使用 HTTPS 建立並提供 Web 應用程式。您可以透過 gulp (在下一節說明) 建立自我主控的 HTTPS 站台，或者也可以使用 Azure。 


### <a name="using-a-self-hosted-https-site"></a>使用自我主控的 HTTPS 站台

gulp-webserver 外掛程式會建立自我主控的 HTTPS 站台。針對所產生的專案，Office 產生器會以名為 serve-static 的工作形式，將這加入 gulpfile.js。使用下列陳述式啟動自我主控的 Web 伺服器︰ 


```
gulp serve-static
```

這會在 https://localhost:8443 啟動 HTTPS 伺服器。


## <a name="develop-your-office-add-in"></a>開發 Office 增益集

您可以使用任何文字編輯器為自訂的 Office 增益集開發檔案。


### <a name="javascript-project-support"></a>JavaScript 專案支援

建立您的專案時，Office 產生器會建立 jsconfig.json 檔案。這個檔案可供您用來推斷專案內的所有 JavaScript 檔案，並讓您不須併入重複的 /// <reference path="../App.js" /> 程式碼區塊。

在 [JavaScript 語言](https://code.visualstudio.com/docs/languages/javascript#_javascript-projects-jsconfigjson) 頁面上深入了解 jsconfig.json 檔案。


### <a name="javascript-intellisense-support"></a>JavaScript Intellisense 支援

此外，即使是撰寫一般 JavaScript，您可以使用 TypeScript 類型定義檔案 (`*.d.ts`) 來提供額外的 IntelliSense 支援。Office 產生器會加入 `tsd.json` 檔案至建立的檔案，其中包含您所選取之專案類型所使用的所有協力廠商程式庫的參考。

使用 Yeoman Office 產生器建立專案之後，您只需要執行下列命令來下載參考的類型定義檔案︰




```
tsd install
```


### <a name="create-a-hello-world-office-add-in"></a>建立 Hello World Office 增益集


本範例中，我們要建立 Hello World 增益集。增益集的 UI 是透過可選擇性地提供 JavaScript 程式設計邏輯的 HTML 檔案提供。 


### <a name="to-create-the-files-for-a-hello-world-add-in"></a>為 Hello World 增益集建立檔案


- 在您的專案資料夾，移至 _[專案資料夾]/app/home_ (在本範例中，它是 myHelloWorldaddin/app/home)，開啟 home.html，並將現有程式碼以下列程式碼取代，這可提供顯示增益集 UI 所需最基本的一組 HTML 標籤。
    
```HTML
        <!DOCTYPE html>  
      <html> 
        <head> 
           <meta charset="UTF-8" /> 
           <meta http-equiv="X-UA-Compatible" content="IE=Edge"/> 
           <link rel="stylesheet" type="text/css" href="program.css" />
         </head> 
   
        <body> 
           <p>Hello World!</p> 
        </body> 
      
       </html> 
```

  
    
- 接下來，在相同的資料夾中，開啟 home.css 檔案，並加入下列 CSS 程式碼。
    
```css
     body 
   { 
        position:relative; 
   } 
   li :hover 
   { 
        text-decoration: underline; 
        cursor:pointer; 
   } 
   h1,h3,h4,p,a,li 
   { 
        font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif; 
        text-decoration-color:#4ec724; 
   } 
```
    
- 然後，回到父專案資料夾，並確定名為 manifest-myHelloWorldaddin.xml 的 XML 檔案包含下列 XML 程式碼。
    
     >**重要事項**  `<id>` 標記中的值是 yeoman 產生器產生專案時所建立的 GUID。請勿變更 Yeoman 產生器為您增益集所建立的 GUID。如果主應用程式是 Azure，`SourceLocation` 值會是類似 _https:// [name-of-your-web-app].azurewebsites.net/[path-to-add-in]_ 的 URL。如果您使用自我主控的選項，如下列範例所示，就會是 _https://localhost:8443 / [path-to-add-in]_。

```XML
     <?xml version="1.0" encoding="utf-8"?> 
   <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
              xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
              xsi:type="TaskPaneApp"> 
   <Id>[GUID-for-your-add-in]</Id> 
   <Version>1.0</Version> 
   <ProviderName>Microsoft</ProviderName> 
   <DefaultLocale>EN-US</DefaultLocale> 
   <DisplayName DefaultValue="myHelloWorldaddin"/> 
   <Description DefaultValue="My first app."/> 
    
   <Hosts> 
     <Host Name="Document"/> 
     <Host Name="Workbook"/> 
   </Hosts>
    
   <DefaultSettings> 
     <SourceLocation DefaultValue="https://localhost:8443/app/home/home.html"/> 
   </DefaultSettings> 
   
   <Permissions>ReadWriteDocument</Permissions>
    
   </OfficeApp> 
```


### <a name="running-the-add-in-locally"></a>在本機執行增益集


若要在本機測試增益集，請開啟您的瀏覽器，並輸入您的 home.html 檔案的 URL。這可以是在 Web 伺服器或自我主控的 HTTPS 站台上。如果您在本機主控它，只需在您的瀏覽器輸入該 URL。在我們的範例中，它是 `https://localhost:8443/app/home/home.html`。 

您會看到錯誤訊息說明「此網站的安全性憑證有問題」。選取 [繼續瀏覽此網站...]，然後您會看到文字 "Hello World!"


 >**附註：**產生的增益集隨附自我簽署的憑證及金鑰；請將它們新增到您的憑證信任授權單位清單，使得瀏覽器不會發出憑證警告。如果您想要使用自己的自我簽署憑證，請參閱 [gulp-webserver](https://www.npmjs.com/package/gulp-webserver) 文件。如需如何信任 OS X Yosemite 中的憑證的指示，請參閱[此 KB 文章 #PH18677](https://support.apple.com/kb/PH18677?locale=en_US)。


## <a name="install-the-add-in-for-testing"></a>安裝增益集進行測試

您可以使用側載來安裝增益集以進行測試︰


- [側載 Office 增益集來進行測試](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [側載 Outlook 增益集來進行測試](../outlook/testing-and-tips.md)
    
或者，您可以發佈增益集至目錄或網路共用，並以一般使用者的方式安裝它。如需詳細資訊，請參閱[建立工作窗格和內容增益集的網路共用資料夾目錄](https://technet.microsoft.com/en-us/browser/fp123503(v=office.14))。


## <a name="debugging-your-office-add-in"></a>偵錯 Office 增益集

偵錯增益集有不同的方式︰


- 您可以使用 Office Web 用戶端，然後開啟瀏覽器的開發人員工具並偵錯增益集，就像任何其他用戶端 JavaScript 應用程式一樣。 
    
- 如果您在 Windows 10 上使用桌面版 Office，您可以[在 Windows 10 上使用 F12 開發人員工具偵錯增益集](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)。
    



## <a name="additional-resources"></a>其他資源



- [在 Visual Studio 中建立和偵錯 Office 增益集](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
    
