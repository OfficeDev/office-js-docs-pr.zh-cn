---
title: Office.context.mailbox - 要求集 1.2
description: Outlook邮箱 API 要求集 1.2 版本的邮箱对象模型。
ms.date: 03/18/2020
ms.localizationpriority: medium
---

# <a name="mailbox-requirement-set-12"></a>邮箱 (要求集 1.2) 

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../outlook/understanding-outlook-add-in-permissions.md)| 受限|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

## <a name="properties"></a>属性

| 属性 | 最小值<br>权限级别 | 模式 | 返回类型 | 最小值<br>要求集 |
|---|---|---|---|:---:|
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-diagnostics-member) | ReadItem | 撰写<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-ewsurl-member) | ReadItem | 撰写<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>Read | [项目](/javascript/api/outlook/office.item?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-userprofile-member) | ReadItem | 撰写<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>方法

| 方法 | 最小值<br>权限级别 | 模式 | 最小值<br>要求集 |
|---|---|---|:---:|
| [convertToLocalClientTime (timeValue) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-converttolocalclienttime-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToUtcClientTime (输入) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-converttoutcclienttime-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-displayappointmentform-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-displaymessageform-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-displaynewappointmentform-member(1)) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(1)) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-getuseridentitytokenasync-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#outlook-office-mailbox-makeewsrequestasync-member(1)) | ReadWriteMailbox | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
