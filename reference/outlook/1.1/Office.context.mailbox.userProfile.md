

# <a name="userprofile"></a>userProfile

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

### <a name="members"></a>成員

####  <a name="displayname-:string"></a>displayName：字串

取得使用者的顯示名稱

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-:string"></a>emailAddress：字串

取得使用者的 SMTP 電子郵件地址。

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-:string"></a>timeZone︰字串

取得使用者的預設時區。

##### <a name="type:"></a>類型：

*   字串

##### <a name="requirements"></a>需求

|需求| 值|
|---|---|
|[最低信箱需求集合版本](../tutorial-api-requirement-sets.md)| 1.0|
|[最低權限等級](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用的 Outlook 模式| 撰寫或讀取|

##### <a name="example"></a>範例

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```