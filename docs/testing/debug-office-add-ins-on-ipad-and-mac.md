
# 在 iPad 和 Mac 上偵錯 Office 增益集

您可以在 Windows 上使用 Visual Studio 來開發並偵錯增益集，但您無法使用它來偵錯 iPad 或 Mac 上的增益集。由於增益集是使用 HTML 和 Javascript 開發，因此設計為跨平台使用，但不同瀏覽器呈現 HTML 的方式可能有細微差異。本文說明如何偵錯 iPad 或 Mac 上執行的增益集。 

## 以 Vorlon.js 偵錯 

Vorlon.js 是網頁的偵錯工具，類似於 F12 工具，設計用來遠端工作，並可讓您跨不同的裝置偵錯網頁。如需詳細資訊，請參閱 [Vorlon 網站](http://www.vorlonjs.com)。  

若要安裝和設定 Vorlon： 

1.  安裝 [Node.js](https://nodejs.org) (如果尚未安裝)。 

2.  搭配使用下列命令與 npm 來安裝 Vorlon︰`sudo npm i -g vorlon` 

3.  使用命令 `vorlon` 執行 Vorlon 伺服器。 

4.  開啟瀏覽器視窗並移至 [http://localhost:1337](http://localhost:1337)，也就是 Vorlon 介面。

5.  將下列指令碼標記新增至增益集的 Home.html 檔 (或主要 HTML 檔案) 的 `<head>` 區段︰
```    
<script src="http://localhost:1337/vorlon.js"></script>    
```  

>**附註︰**您必須啟用 Vorlon 中的 HTTPS 才可使用 Vorlon.js 來偵錯增益集。 若要了解如何執行這項操作，請參閱[偵錯 Office 增益集的 VorlonJS 外掛程式](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/)。

現在，每當在裝置上開啟增益集，它就會顯示在 Vorlon的用戶端清單中 (位於 Vorlon 介面的左邊)。您可以從遠端反白顯示 DOM 元素、從遠端執行命令，以及其他動作等等。  

![顯示 Vorlon.js 介面的螢幕擷取畫面](../../images/vorlon_interface.png)

Office 增益集的專用 Vorlon 外掛程式新增額外的功能，例如與 Office.js API 互動。如需詳細資訊，請參閱部落格貼文 [用於偵錯 Office 增益集的 VorlonJS 外掛程式](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/)。若要啟用 Office 增益集外掛程式︰ 

1.  使用下列命令，在本機複製 Vorlon.js GitHub 儲存機制的裝置分支︰ 
```
git clone https://github.com/MicrosoftDX/Vorlonjs.git
git checkout dev
npm install
```

2.  開啟 /Vorlon/Server/config.json 中的 **config.json** 檔。若要啟動 Office 增益集外掛程式，請將 **enabled** 屬性設為 **true**。

![顯示 config.json 的外掛程式區段的螢幕擷取畫面](../../images/vorlon_plugins_config.png) 
