---
title: "\"Context.subname\"-\"邮箱-要求集 1.3\""
description: ''
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 434c89cde3fee6fd742291524c4998ec3af3cebd
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163771"
---
# <a name="mailbox"></a>邮箱

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最低权限级别](../../../outlook/understanding-outlook-add-in-permissions.md)| 受限|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

## <a name="properties"></a>属性

| 属性 | 最低<br>权限级别 | 型号 | 返回类型 | 最低<br>要求集 |
|---|---|---|---|:---:|
| [过程](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#diagnostics) | ReadItem | 撰写<br>读取 | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.3) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#ewsurl) | ReadItem | 撰写<br>读取 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>读取 | [项](/javascript/api/outlook/office.item?view=outlook-js-1.3) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#userprofile) | ReadItem | 撰写<br>读取 | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| 方法 | 最低<br>权限级别 | 型号 | 最低<br>要求集 |
|---|---|---|:---:|
| [convertToEwsId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#converttoewsid-itemid--restversion-) | 受限 | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#converttolocalclienttime-timevalue-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#converttorestid-itemid--restversion-) | 受限 | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#converttoutcclienttime-input-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#displayappointmentform-itemid-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#displaymessageform-itemid-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#displaynewappointmentform-parameters-) | ReadItem | 读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#getcallbacktokenasync-callback--usercontext-) | ReadItem | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
