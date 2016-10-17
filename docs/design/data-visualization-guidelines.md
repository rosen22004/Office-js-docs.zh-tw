
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a>Office 增益集的資料視覺效果樣式指導方針

良好的資料視覺效果可以協助使用者在他們的資料中尋找深入資訊。他們可以使用這些深入資訊訴說通知和勸說的故事。本文提供指導方針以協助您在 Excel 和其他 Office App 的增益集中設計有效的資料視覺效果。

建議您使用 [Office UI Fabric](http://dev.office.com/fabric) 來建立資料視覺效果的組件區塊。Office UI Fabric 包含樣式和元件，可順暢地與 Office 的外觀整合。 

<!--The following figure shows a data visualization in an add-in that uses Fabric.

![Image of a data visualization with Fabric elements applied**](../../images/fabric-data-visualization.png) 

-->

## <a name="data-visualization-elements"></a>資料視覺效果元素

資料視覺效果共用一般架構和通用視覺和互動式元素，包括標題、標籤和資料繪圖，如下圖所示。

![具有標題、座標軸、圖例和繪圖區之折線圖的影像](../../images/data_visualization_line_chart.png)
![具有座標軸、格線、圖例與資料繪圖之直條圖的影像](../../images/data_visualization_column_chart.png)

### <a name="chart-titles"></a>圖表標題

遵循圖表標題的這些指導方針：

- 讓您的圖表標題簡單易讀。放置它們，以建立與圖表其餘部分相對的清楚視覺階層。
- 一般而言，使用句子大寫 (第一個單字大寫)。若要建立對比或要加強階層，您可以全部使用大寫，但全部大寫應該謹慎使用。
- 合併 [Office UI Fabric 類型坡形](http://dev.office.com/fabric#/styles/typography)，讓您的圖表與使用 Segoe 的 Office UI 一致。您也可以使用不同的字樣來區分圖表內容與 UI。
- 使用無襯線字樣來搭配大型計數器。

下列範例顯示在圖表標題中使用的襯線與無襯線字樣。請注意比例對比和如何有效地使用空白字元，建立強式的視覺階層。

![具有襯線字型資料視覺效果的影像](../../images/data_visualization_serif.png)
![具有無襯線字型資料視覺效果的影像](../../images/data_visualization_san_serif.png)

### <a name="axis-labels"></a>座標軸標籤

讓座標軸標籤夠濃以便清楚地閱讀，並讓文字與背景色彩之間有足夠的對比比例。請確定它們不是那麼濃，而與資料筆跡競爭。

淺灰色對於座標軸標籤最有效。如果您使用 Fabric，請參閱[中性色彩調色盤](http://dev.office.com/fabric#/styles/colors)。

### <a name="data-ink"></a>資料筆跡

像素，代表圖表中實際資料參考做為資料筆跡。這應該是視覺效果的中央焦點。避免使用陰影、濃厚的外框，或歪曲或與資料競爭的不必要設計元素。只有當資料值繫結至色彩值時才使用漸層。除非可以測量的目標值繫結至第三個維度，否則請避免立體圖表。

### <a name="color"></a>色彩

選擇遵循作業系統或應用程式佈景主題的色彩，而不是硬式編碼色彩。同時，請確定您所套用的色彩不會讓資料失真。資料視覺效果色彩的誤用可能會導致資料失真的情況，以及不正確的讀取資訊。

如需在資料視覺效果中使用色彩的最佳作法，請參閱下列內容︰


- [彩虹色彩為什麼不是資料視覺效果的最佳選項](http://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [Color Brewer 2.0：製圖的色彩建議](http://colorbrewer2.org/)
- [我想要色調](http://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a>格線

格線通常是精確地讀取圖表所需，但應呈現為次要的視覺化元素，增強資料筆跡，不與其競爭。讓靜態格線細且淺，除非它們特別針對高對比設計。您也可以使用互動，在使用者與圖表互動時，於內容中建立動態、及時顯示的格線。

淺灰色對於格線最有效。如果您使用 Fabric，請參閱[中性色彩調色盤](http://dev.office.com/fabric#/styles/colors)。

下列影像顯示使用格線的資料視覺效果。

![使用格線之資料視覺效果的影像](../../images/data_visualization_gridlines.png)

### <a name="legends"></a>圖例

如有必要，請新增圖例以︰

- 區分數列
- 顯示比例或值的變更

確定您的圖例會增強資料筆跡，而且不會與其競爭。放置圖例：


- 如果所有圖例項目適合放在圖表上方，根據預設清除繪圖區左上方。
- 在繪圖區右上方，如果所有圖例項目不適合放在圖表上方，視需要讓它可捲動。

若要針對可讀性和協助工具最佳化，請將圖例標記對應至相關的圖表圖形。例如，為散佈圖和泡泡圖圖例使用圓形圖例標記。為折線圖使用線段圖例標記。

### <a name="data-labels-and-tooltips"></a>資料標籤和工具提示

確定資料標籤和工具提示有適當的空白字元和類型變化。使用演算法來最小化閉塞以及衝突。例如，工具提示預設可能會出現在資料點右側，但如果偵測到右邊緣則出現在左側。

## <a name="design-principles"></a>設計原則

Office 設計小組已建立下列設計原則集合，我們會在設計 Office 產品套件的新的資料視覺效果時使用這些集合。

## <a name="visual-design-principles"></a>視覺化設計原則


- 視覺效果應該接受和增強資料，使其易於了解。反白顯示資料，只有在必須提供內容時才新增支援的元素。避免不必要的裝飾 (陰影、外框等等)、圖表垃圾或資料扭曲。
- 視覺效果應鼓勵探索，方法是提供豐富的視覺化回饋。使用建立良好的互動模式、介面控制項和明確的系統回饋。
- 實現傳統的設計原則。使用已建立的印刷樣式和視覺化通訊設計原則來增強格式、可讀性和意義。

## <a name="interaction-design-principles"></a>互動設計原則

- 設計以允許探索。
- 允許與會顯示新的深入資訊 (例如，透過拖曳排序) 的物件直接互動。
- 使用簡單、直接、熟悉的互動模型。

如需有關如何設計方便使用的互動式資料視覺效果的詳細資訊，請參閱 [UI 要件及陷阱](http://uitraps.com/)。

## <a name="motion-design-principles"></a>動作設計原則

動作會遵循刺激物。視覺化元素應該以相同的速率朝相同的方向移動。適用於：


- 圖表建立
- 從一個圖表類型轉換為其他圖表類型
- 篩選
- 排序
- 新增或減少資料
- 塗刷或切割資料
- 調整圖表大小

建立原因的認知。演示動畫時︰


- 一次演示一件事。 
- 在變更為資料筆跡之前，演示座標軸的變更。
- 如果物件以相同的速度朝相同的方向移動，則將物件演示和以動畫顯示為群組。
- 以不超過 4-5 個物件的群組演示資料元素。檢視者對於獨立追蹤超過 4-5 個物件有困難。

動作新增意義。


- 動畫會增加使用者對於資料變更的理解、提供內容，並且做為非語言表達註解圖層。
- 動作應該發生在視覺效果的有意義座標空間中。
- 調整視覺動畫。 
- 避免不必要的動畫。

動作需遵循資料。

- 保留資料對應。如果區域繫結到量值，在轉換之間維持該區域。
- 維持一致的動畫設計語言。如果可行的話，將資料視覺效果動畫對應至現有的 Office 動作設計語言。對於類似圖表類型使用類似動畫。

## <a name="accessibility-in-data-visualizations"></a>資料視覺效果中的協助工具

- 請勿使用色彩做為傳達資訊的唯一方法。色盲的人將無法解譯結果。除了色彩之外，還應該儘可能使用形狀、大小與紋理來傳達資訊。
- 讓所有互動式元素 (例如按鈕或挑選清單) 可以從鍵盤存取。
- 傳送協助工具事件至螢幕助讀程式，以宣告焦點變更、工具提示等等。

## <a name="additional-resources"></a>其他資源 

- [資料 + 設計：準備和視覺化資訊的簡單介紹](https://infoactive.co/data-design)
- [建置資料視覺效果的五個最佳程式庫](http://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [量化資訊的視覺顯示](https://www.edwardtufte.com/tufte/books_vdqi)
