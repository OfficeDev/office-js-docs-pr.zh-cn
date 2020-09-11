---
title: "\"Context.subname\"-\"邮箱-要求集 1.2\""
description: Outlook 邮箱 API 要求集的邮箱对象模型的版本为1.2。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: c7be6af6ec65ca2172c2b6290cf3a77522cd9f4a
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431295"
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

| 属性 | 最小值<br>权限级别 | 型号 | 返回类型 | 最小值<br>要求集 |
|---|---|---|---|:---:|
| [过程](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#diagnostics) | ReadItem | 撰写<br>阅读 | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#ewsurl) | ReadItem | 撰写<br>阅读 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>阅读 | [项](/javascript/api/outlook/office.item?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#userprofile) | ReadItem | 撰写<br>阅读 | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| 方法 | 最小值<br>权限级别 | 型号 | 最小值<br>要求集 |
|---|---|---|:---:|
| [convertToLocalClientTime (timeValue) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToUtcClientTime (输入) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#converttoutcclienttime-input-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#displayappointmentform-itemid-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#displaymessageform-itemid-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | 阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | 撰写<br>阅读 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
