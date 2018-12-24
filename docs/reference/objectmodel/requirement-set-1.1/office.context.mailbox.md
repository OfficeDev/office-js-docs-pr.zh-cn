---
title: Office.context.mailbox - 要求集 1.1
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 352132fcc4645463b922cd3bab200fb8efb167b4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433325"
---
# <a name="mailbox"></a>邮箱

### <a name="officeofficemdcontextofficecontextmdmailbox"></a>[Office](Office.md)[.context](Office.context.md).mailbox

为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 受限|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

### <a name="namespaces"></a>命名空间

[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。

[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。

[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。

### <a name="members"></a>成员

#### <a name="ewsurl-string"></a>ewsUrl :String

获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此成员。

远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。

##### <a name="type"></a>类型:

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 阅读|

### <a name="methods"></a>方法

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}

获取包含以本地客户端时间表示的时间信息的字典。

Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。

如果邮件应用程序在 Outlook 中运行，`convertToLocalClientTime` 方法将返回一个值设置为客户端计算机时区的字典对象。如果邮件应用程序在 Outlook Web App 中运行，`convertToLocalClientTime` 方法将返回值设置为 EAC 中指定的时区的字典对象。

##### <a name="parameters"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`timeValue`| 日期|一个 Date 对象|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="returns"></a>返回：

类型：[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

从包含时间信息的字典中获取 Date 对象。

`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)|要转换的本地时间值。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="returns"></a>返回：

包含以 UTC 表示的时间的 Date 对象。

<dl class="param-type">

<dt>
类型</dt>


<dd>日期</dd>

</dl>

####  <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

显示现有日历约会。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。

在 Outlook for Mac 中，您可以使用此方法来显示不属于定期系列的单个约会，或显示定期系列的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。

在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。

如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。

##### <a name="parameters"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`itemId`| 字符串|现有日历约会的 Exchange Web 服务 (EWS) 标识符。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

显示现有邮件。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。

在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。

如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。

不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。

##### <a name="parameters"></a>参数：

|名称| 类型| 描述|
|---|---|---|
|`itemId`| 字符串|现有消息的 Exchange Web 服务 (EWS) 标识符。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

显示用于新建日历约会的表单。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。

在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。

在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。

如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。

##### <a name="parameters"></a>参数：

|名称| 类型| 描述|
|---|---|---|
| `parameters` | 对象 | 描述新约会的参数字典。 |
| `parameters.requiredAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt; | 包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。 |
| `parameters.optionalAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt; | 包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。 |
| `parameters.start` | 日期 | 指定约会的开始日期和时间的 `Date` 对象。 |
| `parameters.end` | Date | 指定约会的结束日期和时间的 `Date` 对象。 |
| `parameters.location` | String | 包含约会位置的字符串。字符串长度限制为最多 255 个字符。 |
| `parameters.resources` | Array.&lt;String&gt; | 包含约会所需资源的字符串数组。数组限制为最多 100 个条目。 |
| `parameters.subject` | String | 包含约会主题的字符串。字符串长度限制为最多 255 个字符。 |
| `parameters.body` | 字符串 | 约会的正文。正文内容限制为最大 32 KB。 |

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 阅读|

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。

`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。

可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。

应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用 `getCallbackTokenAsync` 方法。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| 函数||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。<br/><br/>令牌作为 `asyncResult.value` 属性中的字符串提供。|
|`userContext`| 对象| &lt;可选&gt;|传递给异步方法的任何状态数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 阅读|

##### <a name="example"></a>示例

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

获取用于标识用户和 Office 外接程序的令牌。

`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](https://docs.microsoft.com/outlook/add-ins/authentication)。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| 函数||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。<br/><br/>令牌作为 `asyncResult.value` 属性中的字符串提供。|
|`userContext`| 对象| &lt;可选&gt;|传递给异步方法的任何状态数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="example"></a>示例

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。

> [!NOTE]
> 此方法在下列应用场景不受支持。
> - 在 Outlook for iOS 或 Outlook for Android 中
> - 当加载项载入 Gmail 邮箱中时
> 
> 在这些情况下，加载项应该[使用 REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。

`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。 有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。

你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。

XML 请求必须指定 UTF-8 编码。

```xml
<?xml version="1.0" encoding="utf-8"?>
```

您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。

> [!NOTE]
> 服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。

##### <a name="version-differences"></a>版本差异

当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`data`| 字符串||EWS 请求。|
|`callback`| 函数||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。<br/><br/>EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。 如果结果大小超过 1 MB，则改为返回一条错误消息。|
|`userContext`| 对象| &lt;可选&gt;|传递给异步方法的任何状态数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读|

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