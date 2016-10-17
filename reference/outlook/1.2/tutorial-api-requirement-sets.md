 

# <a name="understanding-api-requirement-sets"></a>了解 API 需求集合

Outlook 增益集使用[資訊清單](https://msdn.microsoft.com/EN-US/library/office/dn592036.aspx)中的 [Requirements](https://msdn.microsoft.com/en-us/library/office/fp123693.aspx) 元素來宣告它們需要的 API 版本。Outlook 增益集一律包含 [Set](https://msdn.microsoft.com/EN-US/library/office/dn592049.aspx) 元素，並且會將 `Name` 屬性設定為 `Mailbox`，將 `MinVersion` 屬性設定為支援增益集案例的最低 API 需求集合。

例如，以下資訊清單片段指出最低需求集合為 1.1：

```
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

所有 Outlook Api 均隸屬於`Mailbox`[需求集合](https://msdn.microsoft.com/EN-US/library/office/dn535871.aspx#SpecifyRequirementSets_intro)。`Mailbox` 需求集合具有版本，而我們發行的每一組新 API 都隸屬於較高版本的集合。並非所有 Outlook 用戶端都支援最新的 API 組合，但如果 Outlook 用戶端宣告支援某個需求集合，它就會支援該需求集合中的所有 API。

在資訊清單中設定最低需求集合版本，能控制增益集顯示的 Outlook 用戶端。如果用戶端不支援的最低需求集合，它就不會載入增益集。例如，如果指定需求集合 1.3 版，這表示增益集將不會顯示在任何不支援至少 1.3 的 Outlook 用戶端中。

## <a name="using-apis-from-later-requirement-sets"></a>使用來自較新需求集合的 API

設定需求集合不會限制增益集可以使用的 API。例如，如果增益集指定需求集合 1.1，但它正在支援 1.3 的 Outlook 用戶端中執行，增益集可以使用來自需求集合 1.3 的 API。\.

若要使用較新的 API，開發人員只要使用標準 JavaScript 技術來檢查 API 是否存在即可

```
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

對於已存在資訊清單指定之需求集合版本中的任何 API 來說，開發人員不需要執行這類檢查。

## <a name="choosing-a-minimum-requirement-set"></a>選擇最低需求集合

在包含開發人員案例所需之關鍵 API 組合的需求集合中，開發人員應使用最舊的需求集合，因為若缺少這些 API，增益集將無法運作。

## <a name="clients"></a>用戶端

下列用戶端支援 Outlook 增益集。

| 用戶端 | 支援的 API 需求集合 |
| --- | --- |
| Outlook 2016 | 1.1, 1.2, 1.3 |
| Mac Outlook 2016 | 1.1 |
| Outlook 2013 | 1.1, 1.2, 1.3 |
| Outlook 網頁版 (Office 365 和 Outlook.com) | 1.1, 1.2, 1.3 |
| Outlook Web App (Exchange 2013 On-Premise) | 1.1 |
| Outlook Web App (Exchange 2016 On-Premise) | 1.1, 1.2. 1.3 |
>**附註** Outlook 2013 中的 1.3 支援已隨著 [2015 年 12 月 8 日的 Outlook 2013 更新 (KB3114349)](https://support.microsoft.com/en-us/kb/3114349) 一同新增
