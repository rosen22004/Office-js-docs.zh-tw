
# 解決 Office 增益集中的相同原始來源原則的限制


瀏覽器強制執行的相同原始來源原則，可防止從一個網域中載入的指令碼取得或操作來自另一個網域的內容。根據預設，這表示要求的 URL 網域必須與目前網頁的網域相同。例如，這個原則會防止一個網域中的網頁對其主控所在位置以外的網域進行 [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) Web 服務呼叫。

因為 Office 增益集主控於瀏覽器控制項中，相同原始來源原則也會套用於其網頁中執行的指令碼。

開發增益集時若要克服相同原始來源原則的強制執行，您可以︰

- 對匿名存取使用 JSON/P。 
    
- 使用權杖型驗證配置來實作伺服器端指令碼。
    
- 使用跨原始來源資源共用 (CORS)。
    
- 使用 IFRAME 和 POST 訊息來建置您自己的 Proxy。
    

## 對匿名存取使用 JSON/P


克服這項限制的方法之一是使用 JSON/P，以為 Web 服務提供Proxy。 您可以藉由包含 `script` 標記加上可指向任何網域上主控之一些指令碼的 `src` 屬性來執行此動作。 您可以以程式設計方式建立 `script` 標記，動態建立供 `src` 屬性指向的 URL，並透過 URI 查詢參數將參數傳遞至該 URL。 Web 服務提供者會在特定 URL 建立和主控 JavaScript 程式碼，並視 URI 查詢參數而定，傳回不同的指令碼。 然後，這些指令碼會在插入的位置執行，並如預期般運作。

以下是 JSON/P 的範例，其會使用可在任何 Office 增益集中運作的技術。

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## 使用權杖型驗證配置來實作伺服器端指令碼


解決相同原始來源原則限制的另一種方法是將增益集網頁實作為 ASP 頁面，其使用 OAuth 或會在 Cookie 中快取憑證。

如需使用 OAuth 進行驗證的範例，請參閱 [Twitter 的 OAuth 的 SharePoint 網頁組件](http://aidangarnish.net/post/Twitter-SharePoint-Web-Part-With-OAuth)。

如需示範如何使用 `System.Net` 中的 `Cookie` 物件來取得及設定 Cookie 值的伺服器端程式碼的範例，請參閱 [Value](http://msdn2.microsoft.com/EN-US/library/4f772twc) 屬性。


## 使用跨原始來源資源共用 (CORS)


如需使用 [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) 的跨原始來源資源共用功能的範例，請參閱 [XMLHttpRequest2 的新祕訣](http://www.html5rocks.com/en/tutorials/file/xhr2/)中的＜跨原始來源資源共用 (CORS)＞一節。


## 使用 IFRAME 和 POST 訊息來建置您自己的 Proxy


如需如何使用 IFRAME 和 POST 訊息來建置您自己的 Proxy 的範例，請參閱[跨視窗訊息](http://ejohn.org/blog/cross-window-messaging/)。


## 其他資源


- [Office 增益集的隱私權和安全性](../../docs/develop/privacy-and-security.md)
    
