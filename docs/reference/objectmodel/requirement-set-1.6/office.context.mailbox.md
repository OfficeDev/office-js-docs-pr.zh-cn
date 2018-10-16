
# <a name="mailbox"></a>mailbox

### [Office](Office.md)[.context](Office.context.md). mailbox

为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="members-and-methods"></a>成员和方法

| 成员 | 类型 |
|--------|------|
| [ewsUrl](#ewsurl-string) | 成员 |
| [restUrl](#resturl-string) | 成员 |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | 方法 |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | 方法 |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | 方法 |
| [convertToRestId](#converttorestiditemid-restversion--string) | 方法 |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | 方法 |
| [displayAppointmentForm](#displayappointmentformitemid) | 方法 |
| [displayMessageForm](#displaymessageformitemid) | 方法 |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | 方法 |
| [displayNewMessageForm](#displaynewmessageformparameters) | 方法 |
| [getCallbackTokenAsync](#getcallbacktokenasyncoptions-callback) | 方法 |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | 方法 |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | 方法 |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | 方法 |

### <a name="namespaces"></a>命名空间

[诊断](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。

[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 加载项中的邮件或约会的方法和属性。

[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户信息。

### <a name="members"></a>成员

#### <a name="ewsurl-string"></a>ewsUrl： 字符串

获取此电子邮件帐户的 Exchange Web 服务 (EWS) 端点 URL。仅限阅读模式。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此成员。

远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。

应用必须在其清单中指定拥有 **ReadItem** 权限，才能在阅读模式中调用 `ewsUrl` 成员。

在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。

##### <a name="type"></a>类型：

*   字符串

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

#### <a name="resturl-string"></a>restUrl :String

获取此电子邮件帐户的 REST 端点 URL。

`restUrl` 值可用于对用户邮箱进行 [REST API](https://docs.microsoft.com/outlook/rest/) 调用。

应用必须在其清单中指定拥有 **ReadItem** 权限，才能在阅读模式中调用 `restUrl` 成员。

在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。

##### <a name="type"></a>类型：

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

### <a name="methods"></a>方法

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

添加支持事件的事件处理程序。

目前，唯一支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该事件类型。此事件由实现可固定任务窗格的加载项使用，并允许加载项根据当前选定的项目刷新任务窗格 UI。

##### <a name="parameters"></a>参数：

| 名称 | 类型 | 属性 | 说明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || 应调用处理程序的事件。 |
| `handler` | Function || 用于处理事件的函数。此函数必须接受单个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。 |
| `options` | Object | &lt;optional&gt; | 包含一个或多个以下属性的对象文本。 |
| `options.asyncContext` | Object | &lt;optional&gt; | 开发人员可以提供他们想要在回调方法中访问的任何对象。 |
| `callback` | 函数| &lt;optional&gt;|方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="example"></a>示例

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

将适用 REST 格式化的项目 ID 转换为 EWS 格式。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

通过 REST API 检索的项 ID（如 [Outlook 邮件 API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](http://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`itemId`| 字符串|适用 Outlook REST API 进行格式化的项目 ID。|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|值指示用于检索项目 ID 的 Outlook REST API 版本。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="returns"></a>返回：

类型：String

##### <a name="example"></a>示例

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}

获取包含以本地客户端时间表示时间信息的字典。

Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。

如果在 Outlook 中运行邮件应用程序，`convertToLocalClientTime` 方法将返回多个值设置为客户端计算机时区的字典对象。如果在 Outlook Web App 中运行邮件应用程序，`convertToLocalClientTime` 方法将返回多个值设置为 EAC 中指定的时区的字典对象。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`timeValue`| Date|一个 Date 对象|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="returns"></a>返回：

返回：LocalClientTime[ ](/javascript/api/outlook_1_6/office.LocalClientTime)

####  <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

将适用 EWS 格式化的项目 ID 转换为 REST 格式。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用与 REST API 不同的格式（例如 [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](http://graph.microsoft.io/)）。`convertToRestId` 方法将适用 EWS 格式化的 ID 转换为正确的 REST 格式。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`itemId`| 字符串|适用于 Exchange Web 服务 (EWS) 进行格式化的项目 ID。|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|值指示转换的 ID 所使用的 Outlook REST API 版本。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restricted|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="returns"></a>返回：

类型：String

##### <a name="example"></a>示例

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

从包含时间信息的字典中获取 Date 对象。

`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)|要转换的本地时间值。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="returns"></a>返回：

以UTC格式表示时间的 Date 对象

<dl class="param-type">

<dt>类型</dt>

<dd>Date</dd>

</dl>

####  <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

显示现有日历约会。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

`displayAppointmentForm` 方法将在桌面新窗口中或移动设备对话框中打开现有的日历约会。

在 Outlook for Mac 中，您可以使用此方法来显示非重复性的单个约会，或显示重复系列中的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。

在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。

如果指定的项目标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`itemId`| String|现有日历约会的 Exchange Web 服务 (EWS) 标识符。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="example"></a>示例

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

显示一封现有邮件。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

在 Outlook Web App 中，`displayMessageForm` 方法将在桌面新窗口中或移动设备对话框中打开一封现有邮件。

在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。

如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。

请勿使用 `displayMessageForm` 配合 `itemId` 表示约会。 使用 `displayAppointmentForm` 方法显示一个现有约会， 并 `displayNewAppointmentForm` 显示一个创建新约会的窗体。

##### <a name="parameters"></a>参数：

|名称| 类型| 说明|
|---|---|---|
|`itemId`| String|现有邮件的 Exchange Web 服务 (EWS) 标识符。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="example"></a>示例

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

显示用于新建日历约会的窗体。

> [!NOTE]
> 在 Outlook for iOS 或 Outlook for Android 中不支持此方法。

`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充参数内容。

在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含**保存**按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。

在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。

如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。

##### <a name="parameters"></a>参数：

> [!NOTE]
> 所有参数都是可选参数。

|名称| 类型| 说明|
|---|---|---|
| `parameters` | Object | 描述新约会的参数字典。 |
| `parameters.requiredAttendees` | 数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt; | 包含电子邮件地址的字符串数组或包含约会的每个必需与会者 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。 |
| `parameters.optionalAttendees` | 数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt; | 包含电子邮件地址的字符串数组或包含约会的每个可选与会者 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。 |
| `parameters.start` | Date | 指定约会开始日期和时间的 `Date` 对象。 |
| `parameters.end` | Date | 指定约会的结束日期和时间的 `Date` 对象。 |
| `parameters.location` | 字符串 | 包含约会位置的字符串。字符串长度限制为最多 255 个字符。 |
| `parameters.resources` | 数组.&lt;字符串&gt; | 包含约会所需资源的字符串数组。数组限制为最多 100 个条目。 |
| `parameters.subject` | 字符串 | 包含约会主题的字符串。字符串长度限制为最多 255 个字符。 |
| `parameters.body` | String | 约会的正文。正文内容限制为最大 32 KB。 |

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 阅读|

##### <a name="example"></a>示例

```
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

#### <a name="displaynewmessageformparameters"></a>displayNewMessageForm(parameters)

显示用于新建邮件的窗体。

`displayNewMessageForm` 方法将打开可让用户新建邮件的窗体。 如果指定了参数，将使用参数的内容自动填充邮件窗体字段。

如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。

##### <a name="parameters"></a>参数：

> [!NOTE]
> 所有参数都是可选参数。

|名称| 类型| 说明|
|---|---|---|
| `parameters` | Object | 描述新邮件的参数字典。 |
| `parameters.toRecipients` | 数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt; | 包含电子邮件地址的字符串数组或包含收件人行上每个收件人的 `EmailAddressDetails` 对象的数组。 数组限制为最多 100 个条目。 |
| `parameters.ccRecipients` | 数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt; | 包含电子邮件地址的字符串数组或包含抄送行上每个收件人的 `EmailAddressDetails` 对象的数组。 数组限制为最多 100 个条目。 |
| `parameters.bccRecipients` | 数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt; | 包含电子邮件地址的字符串数组或包含密件抄送行上每个收件人的 `EmailAddressDetails` 对象的数组。 数组限制为最多 100 个条目。 |
| `parameters.subject` | 字符串 | 包含邮件主题的字符串。 字符串长度限制为最多 255 个字符。 |
| `parameters.htmlBody` | 字符串 | 邮件的 HTML 正文。 正文内容限制为最大 32 KB。 |
| `parameters.attachments` | Array.&lt;Object&gt; | JSON 对象（是文件或项目附件）的数组。 |
| `parameters.attachments.type` | 字符串 | 指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。 |
| `parameters.attachments.name` | 字符串 | 一个包含附件的名称的字符串，最多包含 255 个字符。|
| `parameters.attachments.url` | 字符串 | 仅在 `type` 设置为 `file` 时才使用。文件位置的 URI。 |
| `parameters.attachments.isInline` | Boolean | 仅在 `type` 设置为 `file` 时才使用。如果为 `true`，表示将在邮件正文中嵌入显示附件，并且不应在附件列表中显示。 |
| `parameters.attachments.itemId` | 字符串 | 仅在 `type` 设置为 `item` 时才使用。 你想要附加到新邮件的现有电子邮件的 EWS 项目 id。 字符串最多达 100 个字符。 |


##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低的邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 阅读|

##### <a name="example"></a>示例

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
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

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync([选项] 回调)

获取一个包含用于调用 REST API 或 Exchange Web 服务令牌的字符串。

`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取不透明令牌。回调令牌的生存期为 5 分钟。

> [!NOTE]
> 建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。 

**REST 令牌**

请求 REST 令牌 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非加载项在其清单中指定了 [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。

在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。

**EWS 令牌**

请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。

加载项应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
| `options` | Object | &lt;optional&gt; | 包含一个或多个以下属性的对象文本。 |
| `options.isRest` | Boolean |  &lt;optional&gt; | 确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。 |
| `options.asyncContext` | Object |  &lt;optional&gt; | 传递给异步方法的任何状态数据。 |
|`callback`| function||方法完成后，通过单个参数调用 `callback` 参数中传递的函数， `asyncResult`, 是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象。令牌以 `asyncResult.value` 属性字符串形式提供。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写和阅读|

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(回调, [userContext])

获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。

`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取不透明令牌。回调令牌的生存期为 5 分钟。

可以将令牌和附件标识符或项标识符传递至第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。

应用必须在其清单中指定拥有 **ReadItem** 权限，才能在阅读模式中调用 `getCallbackTokenAsync` 方法。

在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| function||方法完成后，通过单个参数 `asyncResult`（这是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用 `callback` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。|
|`userContext`| Object| &lt;optional&gt;|传递给异步方法的任何状态数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写和阅读|

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

获取用于标识用户和 Office 加载项的令牌。

`getUserIdentityTokenAsync` 方法返回可以用于标识和[在第三方系统上验证加载项和用户](https://docs.microsoft.com/outlook/add-ins/authentication)的令牌。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`callback`| function||方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)对象）调用在 `callback` 参数中传递的函数。<br/><br/>令牌作为 `asyncResult.value` 属性中的字符串提供。|
|`userContext`| Object| &lt;optional&gt;|传递给异步方法的任何状态数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

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
> 在以下方案中不支持此方法。
> - 在 Outlook for iOS 或 Outlook for Android 中
> - 在 Gmail 邮箱中加载加载项时
> 
> 在这些情况下, 加载项应转而[使用 REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) 访问用户的邮箱。

`makeEwsRequestAsync` 方法代表加载项向 Exchange 发送 EWS 请求。 有关支持的 EWS 操作列表，请参阅[从 Outlook 加载项调用 Web 服务](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) 。

不能通过 `makeEwsRequestAsync` 方法请求与“文件夹”关联的项。

XML 请求必须指定 UTF-8 编码。

```
<?xml version="1.0" encoding="utf-8"?>
```

加载项必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。欲知使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定邮件加载项访问用户邮箱的权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。

> [!NOTE]
> 服务器管理员必须在 Client Access Server EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。

##### <a name="version-differences"></a>版本差异

当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法时，应当将编码值设置为 `ISO-8859-1`。

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

当邮件应用在 Outlook 网页版中运行时，不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定邮件应用是正在 Outlook 中运行还是在 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的 Outlook 版本。

##### <a name="parameters"></a>参数：

|名称| 类型| 属性| 说明|
|---|---|---|---|
|`data`| 字符串||EWS 请求。|
|`callback`| function||方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。<br/><br/>EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。 如果结果的大小超过 1 MB，将转而返回一条错误消息。|
|`userContext`| Object| &lt;optional&gt;|传递给异步方法的任何状态数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低邮箱要求集版本](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[适用的 Outlook 模式](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 撰写或阅读​|

##### <a name="example"></a>示例

下面的示例调用 `makeEwsRequestAsync`  以使用  `GetItem` 操作获取项目的主题。

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