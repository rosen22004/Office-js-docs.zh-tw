# <a name="define-add-in-commands-in-your-manifest"></a>在資訊清單中定義增益集命令

增益集命令會提供簡單的方法來自訂預設 Office UI，其具有可執行動作的 UI 元素；例如，您可以在功能區上新增自訂按鈕。若要建立命令，您可以將 **[VersionOverrides](../../../reference/manifest/versionoverrides.md)** 節點新增至現有的資訊清單。 

當資訊清單含有 **VersionOverrides** 元素、支援增益集命令的 Word、Excel、Outlook 和 PowerPoint 版本會使用該元素內的資訊以載入增益集。不支援增益集命令的舊版 Office 產品將會忽略元素。

當用戶端應用程式辨識 **VersionOverrides** 節點時，增益集名稱會出現在功能區，不會出現在工作窗格或閱讀/撰寫窗格中。增益集不會出現在這兩個位置中。
 
## <a name="versionoverrides"></a>VersionOverrides

[VersionOverrides](../../../reference/manifest/versionoverrides.md) 元素是根元素，包含增益集實作的增益集命令的資訊。在資訊清單結構描述 v1.1 版及更新版本中提供支援。

**VersionOverrides** 結構描述有兩個版本。

| 結構描述版本 | 描述 |
|----------------|-------------|
| 1.0 | 支援適用於 Office 應用程式桌上型電腦版本的增益集命令。 | 
| 1.1 | 新增[可釘選的工作窗格](./pinnable-taskpane.md)和行動增益集的支援。**附註：**目前只有 Windows 版 Outlook 2016 和 iOS 版 Outlook 提供支援。 |

增益集可以巢狀方式將新版放入舊版內，以支援多個版本的 **VersionOverrides** 結構描述。這可讓用戶端支援較新的版本，以利用新功能，同時還可讓舊版用戶端載入較舊的版本。如需詳細資訊，請參閱[實作多個版本](../../../reference/manifest/versionoverrides.md#implementing-multiple-versions)。

**VersionOverrides** 元素包括下列子元素：

- [Description](../../../reference/manifest/description.md)
- [Requirements](../../../reference/manifest/requirements.md)
- [Hosts](../../../reference/manifest/hosts.md)
- [Resources](../../../reference/manifest/resources.md)
- [VersionOverrides](../../../reference/manifest/versionoverrides.md)

下圖顯示用來定義增益集命令的元素階層。 

![資訊清單中增益集命令元素的階層](../../../images/080da303-51c4-4882-b74a-7ba11517c0ad.png)

## <a name="sample-manifests"></a>範例資訊清單

如需實作 Word、Excel 和 PowerPoint 的增益集命令的範例資訊清單，請參閱[簡單增益集命令範例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple)。

如需實作 Outlook 的增益集命令的範例資訊清單，請參閱 [Outlook 增益集的簡單資訊清單檔案](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml)。

## <a name="additional-resources"></a>其他資源

- [Outlook 的增益集命令](../../outlook/add-in-commands-for-outlook.md)
    
- [Outlook 增益集資訊清單](../../outlook/manifests/manifests.md)
    
- [Outlook 增益集命令示範範例](https://github.com/jasonjoh/command-demo)
