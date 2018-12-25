---
title: Office.context.mailbox.userProfile - 要求集 1.4
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 55d0a789c8e46fd3f6ee69f39cf33f7e7d94c322
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432646"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

### <a name="members"></a>成员

####  <a name="displayname-string"></a>displayName :String

获取用户的显示名称。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :String

获取用户的 SMTP 电子邮件地址。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :String

获取用户的默认时区。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```