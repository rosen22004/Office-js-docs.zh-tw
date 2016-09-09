

# Office 物件
代表增益集的執行個體，其可供存取 API 的最上層物件。

|||
|:-----|:-----|
|**主應用程式︰**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**上次變更於**|1.1|

```js
Office
```


## 成員


**屬性**

|||
|:-----|:-----|
|名稱|說明|
|[內容](../../reference/shared/office.context.md)|取得 Context 物件，其代表增益集的執行階段環境，並可供存取 API 的最上層物件。|
|[cast.item](../../reference/shared/office.cast.item.md)|在 Visual Studio 中提供撰寫或讀取模式的訊息和約會所特有的 IntelliSense。 <br/><br/><blockquote>**附註**  僅適用於在 Visual Studio 中開發 Outlook 增益集時的設計階段。 </blockquote>|

**方法**

|||
|:-----|:-----|
|名稱|說明|
|[select](../../reference/shared/office.select.md)|建立承諾，以依據傳入的選取器字串傳回繫結。|
|[useShortNamespace](../../reference/shared/office.useshortnamespace.md)|切換完整的 **Microsoft.Office.WebExtension** 命名空間之 **Office** 別名的開和關。|

**事件**

|||
|:-----|:-----|
|名稱|說明|
|[初始化](../../reference/shared/office.initialize.md)|已載入執行階段環境，且增益集準備好開始與裝載它的文件互動時，就會發生。|

## 備註

**Office** 物件可讓開發人員針對 Initialize 事件實作回呼函數，並可用於存取 [Context](../../reference/shared/context.md) 物件。


## 支援詳細資料


下列矩陣中的大寫 Y，表示在相對應的 Office 主應用程式中支援此物件。空白儲存格表示 Office 主應用程式不支援此物件。

如需有關 Office 主應用程式與伺服器需求的詳細資訊，請參閱[執行 Office 增益集的需求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows desktop**|**Office Online (在瀏覽器中)**|**Office for iPad**|**裝置適用的 OWA**|**Mac 版 Outlook**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|是|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**增益集類型**|內容、Outlook、工作窗格|
|**文件庫**|Office.js|
|**命名空間**|Office|

## 支援歷程記錄


|**版本**|**變更**|
|:-----|:-----|
|1.1|新增 iPad 版 Office 中對 Excel、PowerPoint 和 Word 的支援。|
|1.1|<ul><li>針對 <a href="6c4b2c16-d4fb-4ecf-b72c-1e33b205daaf.htm">context</a>，新增支援在 Access 的內容增益集中取得執行階段內容。</p></li><li><p>針對 <a href="23aeb136-da1f-4127-a798-99dc27bc4dae.htm">select</a>，新增支援在 Access 的內容增益集中選取資料表繫結。</li><li>針對 <a href="9a4d5c7d-fcc4-4e8f-bef2-f2a8d8b4ae00.htm">useShortNamespace</a>，新增支援 Access 的內容增益集。</li><li>針對 <a href="727adf79-a0b5-48d2-99c7-6642c2c334fc.htm">initialize</a>，新增支援在 Access 的內容增益集中初始化。</li></ul>|
|1.0|已導入|

