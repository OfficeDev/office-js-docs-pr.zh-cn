---
title: Office。上下文要求集1。4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9103f6666b9a7903393729bceb157dfae615da06
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163750"
---
# <a name="context"></a>context

### <a name="officecontext"></a>[Office](office.md).context

在所有 Office 应用中，上下文提供外接程序使用的共享接口。 此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.4)"。

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="properties"></a>属性

| 属性 | 型号 | 返回类型 | 最低<br>要求集 |
|---|---|---|:---:|
| [contentLanguage](#contentlanguage-string) | 撰写<br>读取 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [过程](#diagnostics-contextinformation) | 撰写<br>读取 | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | 撰写<br>读取 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [host](#host-hosttype) | 撰写<br>读取 | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [mailbox](office.context.mailbox.md) | 撰写<br>读取 | [邮箱](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [平台](#platform-platformtype) | 撰写<br>读取 | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [满足](#requirements-requirementsetsupport) | 撰写<br>读取 | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | 撰写<br>读取 | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | 撰写<br>读取 | [UI](/javascript/api/office/office.ui?view=outlook-js-1.4) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>属性详细信息

#### <a name="contentlanguage-string"></a>contentLanguage： String

获取用户指定的用于编辑项的区域设置（语言）。

此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。

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

#### <a name="diagnostics-contextinformation"></a>诊断： [ContextInformation](/javascript/api/office/office.contextinformation)

获取有关加载项在其中运行的环境的信息。

##### <a name="type"></a>类型

*   [ContextInformation](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage： String

获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。

`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。

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

#### <a name="host-hosttype"></a>主机： [HostType](/javascript/api/office/office.hosttype)

获取运行外接程序的 Office 应用程序主机。

##### <a name="type"></a>类型

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a>platform： [PlatformType](/javascript/api/office/office.platformtype)

提供在其上运行外接的平台。

##### <a name="type"></a>类型

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a>要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

提供用于确定当前主机和平台上支持的要求集的方法。

##### <a name="type"></a>类型

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a>Requirements

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

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)

获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。

`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。

##### <a name="type"></a>类型

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../outlook/understanding-outlook-add-in-permissions.md)| 受限|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

<br>

---
---

#### <a name="ui-ui"></a>ui： [ui](/javascript/api/office/office.ui)

提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。

##### <a name="type"></a>类型

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|
