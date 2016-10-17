# <a name="officetab-element"></a>OfficeTab 元素
定義您的增益集命令在上面顯示的功能區索引標籤。這可位於預設索引標籤 (不論是 [家用]、[訊息] 或 [會議])，或增益集所定義的自訂索引標籤。這個元素是必要的。

## <a name="child-elements"></a>子元素
|  元素 |  必要  |  描述  |
|:-----|:-----|:-----|
|  群組      | 是 |  定義命令群組。您可以將各個增益集一個群組新增至預設索引標籤。  |


以下是根據主應用程式的有效索引標籤 `id` 值。**粗體**表示的值同時在桌面和線上支援 (例如，Word 2016 for Windows 和 Word Online)。 

### <a name="outlook"></a>Outlook 
- **TabDefault**

### <a name="word"></a>Word
- **TabHome**
- **TabInsert**
- TabWordDesign
- **TabPageLayoutWord**
- TabReferences
- TabMailings
- TabReviewWord
- **TabView**
- TabDeveloper
- TabAddIns
- TabBlogPost
- TabBlogInsert
- TabPrintPreview
- TabOutlining
- TabConflicts
- TabBackgroundRemoval
- TabBroadcastPresentation

### <a name="excel"></a>Excel
- **TabHome**
- **TabInsert**
- TabPageLayoutExcel
- TabFormulas
- **TabData**
- **TabReview**
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabBackgroundRemoval 

### <a name="powerpoint"></a>PowerPoint
- **TabHome**
- **TabInsert**
- **TabDesign**
- **TabTransitions**
- **TabAnimations**
- TabSlideShow
- TabReview
- **TabView**
- TabDeveloper
- TabAddIns
- TabPrintPreview
- TabMerge
- TabGrayscale
- TabBlackAndWhite
- TabBroadcastPresentation
- TabSlideMaster
- TabHandoutMaster
- TabNotesMaster
- TabBackgroundRemoval
- TabSlideMasterHome

### <a name="onenote"></a>OneNote
- **TabHome**
- **TabInsert**
- **TabView**
- TabDeveloper
- TabAddIns

## <a name="group"></a>群組
索引標籤中的一群使用者介面擴充點。一個群組可以有最多六個控制項。**id** 屬性是必要的，且每個 **id** 在資訊清單內必須是唯一的。**id** 是最多為 125 個字元的字串。請參閱[群組元素](./group.md)。

## <a name="officetab-example"></a>OfficeTab 範例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
