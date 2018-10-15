
# <a name="userprofile"></a>userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="members-and-methods"></a>成员和方法

| 成员 | 类型 |
|--------|------|
| [accountType](#accounttype-string) | 成员 |
| [displayName](#displayname-string) | 成员 |
| [emailAddress](#emailaddress-string) | 成员 |
| [timeZone](#timezone-string) | 成员 |

### <a name="members"></a>成员

####  <a name="accounttype-string"></a>accountType: String

> [!NOTE]
> 当前仅 Outlook 2016  for Mac 或更高版本（内部版本 16.9.1212 或更高版本）支持此成员。

获取与邮箱关联用户的帐户类型。下表列出了可能的值。

| 值 | 说明 |
|-------|-------------|
| `enterprise` | 邮箱位于本地 Exchange 服务器上。 |
| `gmail` | 邮箱与 Gmail 帐户关联。 |
| `office365` | 邮箱与 Office 365 工作或学校帐户关联。 |
| `outlookCom` | 邮箱与个人 Outlook.com 帐户关联。 |

##### <a name="type"></a>类型：

*   字符串

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a>displayName :字符串

获取用户的显示名称。

##### <a name="type"></a>类型：

*   字符串

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :字符串

获取用户的 SMTP 电子邮件地址。

##### <a name="type"></a>类型：

*   字符串

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :字符串

获取用户的默认时区。

##### <a name="type"></a>类型：

*   字符串

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```