
# <a name="allowsnapshot-element"></a>AllowSnapshot 項目
指定是否要將內容增益集的快照影像與主文件一起儲存。

 **增益集類型︰**內容


## <a name="syntax:"></a>語法：


```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```


## <a name="contained-in:"></a>內含於：

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="remarks"></a>備註


 **安全性提示：** **AllowSnapshot** 在預設狀況下為 **True**。這會讓使用者在不支援 Office 增益集的主應用程式版本中開啟文件時，可以看到增益集的影像；或在主應用程式無法連線至裝載增益集的伺服器時，提供增益集的靜態影像。不過，這也表示增益集中所顯示可能的機密資訊，可以直接從裝載增益集的文件中存取。

