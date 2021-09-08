---
title: Office.context - 要求集 1.9
description: Office。适用于使用邮箱 API Outlook集 1.9 的外接程序的上下文对象成员。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: f45eec7ce638f4bbb97ad4be9f2ba089905c631d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939055"
---
# <a name="context-mailbox-requirement-set-19"></a>context (Mailbox requirement set 1.9) 

### <a name="officecontext"></a>[Office](office.md).context

Office.context 提供了外接程序在所有应用程序中使用的共享Office接口。 此列表仅记录加载项Outlook接口。有关 Office.context 命名空间的完整列表，请参阅通用 API 中的[Office.context 引用](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

## <a name="properties"></a>属性

| 属性 | 模式 | 返回类型 | 最小值<br>要求集 |
|---|---|---|:---:|
| [auth](#auth-auth) | 撰写<br>读取 | [Auth](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [IdentityAPI 1.3](../../requirement-sets/identity-api-requirement-sets.md) |
| [contentLanguage](#contentlanguage-string) | 撰写<br>读取 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [diagnostics](#diagnostics-contextinformation) | 撰写<br>读取 | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | 撰写<br>读取 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [host](#host-hosttype) | 撰写<br>读取 | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [mailbox](office.context.mailbox.md) | 撰写<br>读取 | [邮箱](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [平台](#platform-platformtype) | 撰写<br>读取 | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [requirements](#requirements-requirementsetsupport) | 撰写<br>读取 | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | 撰写<br>读取 | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | 撰写<br>读取 | [UI](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>属性详细信息

#### <a name="auth-auth"></a>身份验证 [：Auth](/javascript/api/office/office.auth)

通过提供允许 Office 应用程序获取对外接程序 Web 应用程序的访问令牌的方法 ([SSO](../../../outlook/authenticate-a-user-with-an-sso-token.md)) 支持单一登录。 这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。 请参阅 [IdentityAPI 1.3 要求集](../../requirement-sets/identity-api-requirement-sets.md)。

##### <a name="type"></a>类型

*   [Auth](/javascript/api/office/office.auth)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 不适用|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a>contentLanguage： String

获取用户 (编辑) 区域设置语言设置。

该值 `contentLanguage` 反映当前在客户端 **应用程序中** 由 File **> Options > Language** 指定的Office语言设置。

##### <a name="type"></a>类型

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a>diagnostics： [ContextInformation](/javascript/api/office/office.contextinformation)

获取加载项运行环境的信息。

##### <a name="type"></a>类型

*   [ContextInformation](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage：String

获取区域设置 (语言) RFC 1766 语言标记格式，该标记格式由用户为 Office 客户端应用程序的 UI 指定。

该值 `displayLanguage` 反映当前显示 **语言** 设置，该设置 **由**> 客户端应用程序中>选项Office语言。

##### <a name="type"></a>类型

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a>host： [HostType](/javascript/api/office/office.hosttype)

获取Office加载项的应用程序。

> [!NOTE]
> 或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取平台。

##### <a name="type"></a>类型

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a>platform： [PlatformType](/javascript/api/office/office.platformtype)

提供运行加载项的平台。

> [!NOTE]
> 或者，您可以使用[Office.context.diagnostics](#diagnostics-contextinformation)属性获取平台。

##### <a name="type"></a>类型

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a>requirements： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

提供用于确定当前应用程序和平台上支持哪些要求集的方法。

##### <a name="type"></a>类型

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a>[roamingSettings：RoamingSettings](/javascript/api/outlook/office.roamingsettings)

获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。

该对象允许您存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时可以使用该数据 `RoamingSettings` 。

##### <a name="type"></a>类型

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../outlook/understanding-outlook-add-in-permissions.md)| 受限|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

<br>

---
---

#### <a name="ui-ui"></a>[ui：UI](/javascript/api/office/office.ui)

提供可用于在加载项中创建和操作 UI 组件（如对话框）Office方法。

##### <a name="type"></a>类型

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|
