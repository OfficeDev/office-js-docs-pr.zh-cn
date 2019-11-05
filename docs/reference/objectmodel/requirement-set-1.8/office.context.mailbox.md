---
title: "\"Context.subname\"-\"邮箱-要求集 1.8\""
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 3f6d639cdf8bdff6f2df365622f58eba1c4b38e0
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902140"
---
# <a name="mailbox"></a>邮箱

### <a name="officeofficemdcontextofficecontextmdmailbox"></a>[Office](Office.md)[.context](Office.context.md).mailbox

为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| 受限|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="members-and-methods"></a>成员和方法

| 成员 | 类型 |
|--------|------|
| [ewsUrl](#ewsurl-string) | 成员 |
| [masterCategories](#mastercategories-mastercategories) | 成员 |
| [restUrl](#resturl-string) | 成员 |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | 方法 |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | 方法 |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | 方法 |
| [convertToRestId](#converttorestiditemid-restversion--string) | 方法 |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | 方法 |
| [displayAppointmentForm](#displayappointmentformitemid) | 方法 |
| [displayMessageForm](#displaymessageformitemid) | 方法 |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | 方法 |
| [Office.context.mailbox.displaynewmessageform](#displaynewmessageformparameters) | 方法 |
| [getCallbackTokenAsync](#getcallbacktokenasyncoptions-callback) | 方法 |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | 方法 |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | 方法 |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | 方法 |
| [removeHandlerAsync](#removehandlerasynceventtype-options-callback) | 方法 |

### <a name="namespaces"></a>命名空间

[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。

[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。

[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。

### <a name="members"></a>Members

#### <a name="ewsurl-string"></a>ewsUrl：String

获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。

> [!NOTE]
> iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。

远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。

应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。

在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。

##### <a name="type"></a>类型

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategoriesviewoutlook-js-18"></a>masterCategories： [masterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)

获取一个对象，该对象提供用于管理此邮箱上的类别主列表的方法。

> [!NOTE]
> iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。

##### <a name="type"></a>类型

*   [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.8 |
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox |
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读 |

##### <a name="example"></a>示例

本示例获取此邮箱的类别主列表。

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a>restUrl：String

获取此电子邮件帐户的 REST 终结点的 URL。

`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。

应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。

在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。

##### <a name="type"></a>类型

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

### <a name="methods"></a>方法

#### <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

添加支持事件的事件处理程序。

目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。

##### <a name="parameters"></a>Parameters

| 名称 | 类型 | 属性 | 说明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || 应调用处理程序的事件。 |
| `handler` | 函数 || 用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。 |
| `options` | Object | &lt;可选&gt; | 包含一个或多个以下属性的对象文本。 |
| `options.asyncContext` | 对象 | &lt;可选&gt; | 开发人员可以提供他们想要在回调方法中访问的任何对象。 |
| `callback` | 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

将项目 ID 格式化（从 REST 转换为 EWS 格式）。

> [!NOTE]
> iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。

通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。

##### <a name="parameters"></a>Parameters

|名称| 类型| 说明|
|---|---|---|
|`itemId`| 字符串|Outlook REST API 的格式化的项目 ID。|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|指示用于检索项目 ID 的 Outlook REST API 的版本。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| 受限|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="returns"></a>返回：

类型：字符串

##### <a name="example"></a>示例

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-18"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}

获取包含以本地客户端时间表示的时间信息的字典。

Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。

如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。

##### <a name="parameters"></a>Parameters

|名称| 类型| 描述|
|---|---|---|
|`timeValue`| 日期|一个 Date 对象|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="returns"></a>返回：

类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

将项目 ID 格式化（从 EWS 转换为 REST 格式）。

> [!NOTE]
> iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。

与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。

##### <a name="parameters"></a>Parameters

|名称| 类型| 说明|
|---|---|---|
|`itemId`| 字符串|适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|值指示转换的 ID 所使用的 Outlook REST API 的版本。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| 受限|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="returns"></a>返回：

类型：字符串

##### <a name="example"></a>示例

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

从包含时间信息的字典中获取 Date 对象。

`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。

##### <a name="parameters"></a>Parameters

|名称| 类型| 说明|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)|要转换的本地时间值。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="returns"></a>返回：

包含以 UTC 表示的时间的 Date 对象。

键入：日期

##### <a name="example"></a>示例

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

显示现有日历约会。

> [!NOTE]
> iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。

`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。

在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。

在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。

如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。

##### <a name="parameters"></a>参数

|名称| 类型| 说明|
|---|---|---|
|`itemId`| 字符串|现有日历约会的 Exchange Web 服务 (EWS) 标识符。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

显示现有邮件。

> [!NOTE]
> iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。

`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。

在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。

如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。

不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。

##### <a name="parameters"></a>Parameters

|名称| 类型| 说明|
|---|---|---|
|`itemId`| 字符串|现有消息的 Exchange Web 服务 (EWS) 标识符。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

显示用于新建日历约会的表单。

> [!NOTE]
> iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。

`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。

在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”**** 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”**** 按钮。

在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。

如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。

##### <a name="parameters"></a>参数

> [!NOTE]
> 所有参数都是可选的。

|名称| 类型| 说明|
|---|---|---|
| `parameters` | 对象 | 描述新约会的参数字典。 |
| `parameters.requiredAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt; | 包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。 |
| `parameters.optionalAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt; | 包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。 |
| `parameters.start` | 日期 | 指定约会的开始日期和时间的 `Date` 对象。 |
| `parameters.end` | Date | 指定约会的结束日期和时间的 `Date` 对象。 |
| `parameters.location` | 字符串 | 包含约会位置的字符串。字符串长度限制为最多 255 个字符。 |
| `parameters.resources` | Array.&lt;String&gt; | 包含约会所需资源的字符串数组。数组限制为最多 100 个条目。 |
| `parameters.subject` | String | 包含约会主题的字符串。字符串长度限制为最多 255 个字符。 |
| `parameters.body` | 字符串 | 约会的正文。正文内容限制为最大 32 KB。 |

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 阅读|

##### <a name="example"></a>示例

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="displaynewmessageformparameters"></a>Office.context.mailbox.displaynewmessageform （参数）

显示用于创建新邮件的窗体。

`displayNewMessageForm`方法将打开一个窗体，使用户可以创建新邮件。 如果指定了参数，则将使用参数的内容自动填充邮件窗体字段。

如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。

##### <a name="parameters"></a>参数

> [!NOTE]
> 所有参数都是可选的。

|名称| 类型| 说明|
|---|---|---|
| `parameters` | 对象 | 描述新邮件的参数的字典。 |
| `parameters.toRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt; | 包含电子邮件地址的字符串数组，或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。 数组限制为最多 100 个条目。 |
| `parameters.ccRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt; | 包含电子邮件地址的字符串数组，或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。 数组限制为最多 100 个条目。 |
| `parameters.bccRecipients` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt; | 包含电子邮件地址的字符串数组，或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。 数组限制为最多 100 个条目。 |
| `parameters.subject` | 字符串 | 包含邮件主题的字符串。 字符串长度限制为最多 255 个字符。 |
| `parameters.htmlBody` | 字符串 | 邮件的 HTML 正文。 正文内容限制为最大 32 KB。 |
| `parameters.attachments` | Array.&lt;Object&gt; | JSON 对象（是文件或项目附件）的数组。 |
| `parameters.attachments.type` | String | 指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。 |
| `parameters.attachments.name` | 字符串 | 一个包含附件的名称的字符串，最多包含 255 个字符。|
| `parameters.attachments.url` | 字符串 | 仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。 |
| `parameters.attachments.isInline` | 布尔 | 仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。 |
| `parameters.attachments.itemId` | 字符串 | 仅在 `type` 设置为 `item` 时使用。 要附加到新邮件的现有电子邮件的 EWS 项目 id。 字符串最长为 100 个字符。 |


##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 阅读|

##### <a name="example"></a>示例

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync([options], callback)

获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。


            `getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。

> [!NOTE]
> 建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。

在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。

在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。

**REST 令牌**

请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。

在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。

**EWS 令牌**

请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。

外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。

可以将令牌和附件标识符或项标识符传递到第三方系统。 第三方系统使用令牌作为持有者授权令牌，以调用 Exchange Web 服务（EWS） [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作或[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作以检索附件或项目。 例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。

##### <a name="parameters"></a>Parameters

|名称| 类型| 属性| 说明|
|---|---|---|---|
| `options` | 对象 | &lt;可选&gt; | 包含一个或多个以下属性的对象文本。 |
| `options.isRest` | 布尔值 |  &lt;可选&gt; | 确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。 |
| `options.asyncContext` | Object |  &lt;可选&gt; | 传递给异步方法的任何状态数据。 |
|`callback`| 函数||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。<br/><br/>令牌作为 `asyncResult.value` 属性中的字符串提供。<br><br>如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。|

##### <a name="errors"></a>错误

|错误代码|说明|
|------------|-------------|
|`HTTPRequestFailure`|请求失败。 请查看诊断对象，了解 HTTP 错误代码。|
|`InternalServerError`|Exchange 服务器返回了错误。 请查看诊断对象，了解详细信息。|
|`NetworkError`|用户不再连接到网络。 请检查网络连接并重试。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写和阅读|

##### <a name="example"></a>示例

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。

`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。

可以将令牌和附件标识符或项标识符传递到第三方系统。 第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。 例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。

在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。

在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。

##### <a name="parameters"></a>Parameters

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| function||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。<br/><br/>令牌作为 `asyncResult.value` 属性中的字符串提供。<br><br>如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。|
|`userContext`| 对象| &lt;可选&gt;|传递给异步方法的任何状态数据。|

##### <a name="errors"></a>错误

|错误代码|说明|
|------------|-------------|
|`HTTPRequestFailure`|请求失败。 请查看诊断对象，了解 HTTP 错误代码。|
|`InternalServerError`|Exchange 服务器返回了错误。 请查看诊断对象，了解详细信息。|
|`NetworkError`|用户不再连接到网络。 请检查网络连接并重试。|

##### <a name="requirements"></a>要求

|要求|||
|---|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0 | 1.3 |
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem | ReadItem |
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 阅读 | 撰写 |

##### <a name="example"></a>示例

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

获取用于标识用户和 Office 外接程序的令牌。

`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。

##### <a name="parameters"></a>参数

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| function||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。<br/><br/>令牌作为 `asyncResult.value` 属性中的字符串提供。<br><br>如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。|
|`userContext`| 对象| &lt;可选&gt;|传递给异步方法的任何状态数据。|

##### <a name="errors"></a>错误

|错误代码|说明|
|------------|-------------|
|`HTTPRequestFailure`|请求失败。 请查看诊断对象，了解 HTTP 错误代码。|
|`InternalServerError`|Exchange 服务器返回了错误。 请查看诊断对象，了解详细信息。|
|`NetworkError`|用户不再连接到网络。 请检查网络连接并重试。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。

> [!NOTE]
> 此方法在下列应用场景不受支持。
> - 在 iOS 版 Outlook 或 Android 版 Outlook 中
> - 当加载项载入 Gmail 邮箱中时
> 
> 在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。

`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。 有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。

你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。

XML 请求必须指定 UTF-8 编码。

```xml
<?xml version="1.0" encoding="utf-8"?>
```

您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。

> [!NOTE]
> 服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。

##### <a name="version-differences"></a>版本差异

当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。 您可以使用邮箱. hostName 属性确定您的邮件应用程序是在 web 上的 Outlook 中运行还是在桌面客户端上运行。 可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。

##### <a name="parameters"></a>Parameters

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`data`| 字符串||EWS 请求。|
|`callback`| 函数||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。<br/><br/>EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。 如果结果大小超过 1 MB，则改为返回一条错误消息。|
|`userContext`| 对象| &lt;可选&gt;|传递给异步方法的任何状态数据。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a>removeHandlerAsync(eventType, [options], [callback])

删除受支持事件类型的事件处理程序。

目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。

##### <a name="parameters"></a>Parameters

| 名称 | 类型 | 属性 | 说明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || 应撤销处理程序的事件。 |
| `options` | 对象 | &lt;可选&gt; | 包含一个或多个以下属性的对象文本。 |
| `options.asyncContext` | 对象 | &lt;可选&gt; | 开发人员可以提供他们想要在回调方法中访问的任何对象。 |
| `callback` | 函数| &lt;可选&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|