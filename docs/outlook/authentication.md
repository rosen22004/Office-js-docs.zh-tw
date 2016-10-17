
# <a name="authenticate-an-outlook-add-in-by-using-exchange-identity-tokens"></a>使用 Exchange 識別權杖來驗證 Outlook 增益集

Outlook 增益集可以提供客戶來自網際網路上任何位置的資訊，無論是來自主控增益集的伺服器、來自您的內部網路或雲端中某處。不過，如果該資訊受到保護，您的增益集需要一個方式，將 Exchange 電子郵件帳戶與您的資訊服務產生關聯。藉由提供可識別提出要求之電子郵件帳戶的權杖，Exchange 2013 可以為您的增益集啟用單一登入 (SSO)。您可以將這個權杖與您的應用程式已註冊的使用者產生關聯，使得每當增益集連線到您的服務時即可識別使用者。

## <a name="identity-tokens"></a>識別權杖


我們的範例增益集中有兩個使用可公開取得的資訊 - 一個對郵件的地址顯示 Bing 地圖，另一個在郵件中顯示 YouTube 視訊連結的預覽。但您的增益集也可以存取非公開資訊。您可以使用主控增益集的伺服器，將您的增益集連結至內部網路或雲端中某處的資訊。

您可以使用許多不同技術來識別及驗證增益集的使用者。Exchange 2013 藉由為增益集提供可識別特定 Exchange 電子郵件帳戶的識別權杖，簡化使用者驗證。您可以將此權杖與您的服務中已註冊的使用者產生關聯，為使用 Outlook 增益集的客戶啟用單一登入 (SSO)。 

若要在增益集中使用 SSO，程式碼會執行此動作︰


* 呼叫 Outlook 增益集 API 中會傳回識別權杖的函式。
* 將權杖與要求一起傳送到您的伺服器。
* 解壓縮來自伺服器的回應以顯示資訊來自服務的資訊。
    
在伺服器端，事情都更為複雜。當您的伺服器從 Outlook 增益集收到要求時，處理程序的運作方式如下︰

* 伺服器會驗證權杖。您可以使用我們的 [Managed 權杖驗證程式庫](../../docs/outlook/use-the-token-validation-library.md)，或者您也可以為您的服務[建立您自己的程式庫](../../docs/outlook/validate-an-identity-token.md)。
* 伺服器會從權杖查詢唯一識別碼，以了解它是否與已知的識別相關聯。您的服務必須[實作會將識別碼](../../docs/outlook/authenticate-a-user-with-an-identity-token.md)與您的服務的已知使用者比對的方法。
* 如果唯一識別碼比對到先前在伺服器上儲存的一組認證的識別碼，您的伺服器可以以要求的資訊回應，而不需要求您的客戶登入服務。
* 如果唯一識別碼未知，則伺服器會傳送回應，要求使用者使用伺服器的認證登入。
* 如果認證在伺服器上比對到已知的識別，您可以將該識別對應到權杖中的唯一識別碼，以便下一次要求進入時，您的伺服器可以回應，而不需要求其他登入步驟。

 >**附註：**這只是對於如何使用識別權杖的一個建議。一如往常，處理識別和驗證時，必須確定您的程式碼符合組織的安全性需求。

讓我們進入細節。在下列的文章中，我們將使用簡單的 Outlook 增益集，其會將識別權杖和在郵件中找到的電話號碼清單傳送至 Web 服務。 

- [Exchange 識別權杖的內容](../outlook/inside-the-identity-token.md)
- [在 Exchange 中使用識別權杖以從 Outlook 增益集呼叫服務](../outlook/call-a-service-by-using-an-identity-token.md)
- [使用 Exchange 權杖驗證程式庫](../outlvalidate-an-identity-token.md ook/use-the-token-validation-library.md)
- [驗證 Exchange 識別權杖](../outlook/validate-an-identity-token.md )
- [使用 Exchange 的識別權杖來驗證使用者](../outlook/validate-an-identity-token.md)


## <a name="additional-resources"></a>其他資源



- [Outlook 增益集](../outlook/outlook-add-ins.md)
    
- [從 Outlook 增益集呼叫 Web 服務](../outlook/web-services.md)
    


