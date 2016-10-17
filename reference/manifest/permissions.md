
# <a name="permissions-element"></a>Permissions 項目
指定您 Office 增益集的 API 存取層級；您應該根據「最低權限」原則來要求各項權限。

 **增益集類型︰**內容、工作窗格、郵件


## <a name="syntax:"></a>語法：

針對內容和工作窗格增益集：


```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

針對郵件增益集：




```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```


## <a name="contained-in:"></a>內含於：

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## <a name="remarks"></a>備註

如需詳細資訊，請參閱[要求用於內容和工作窗格增益集的 API 權限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)和[了解 Outlook 增益集的權限](../../docs/outlook/understanding-outlook-add-in-permissions.md)。

