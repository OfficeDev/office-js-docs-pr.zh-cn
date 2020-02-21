---
title: "\"Context.subname\"-\"邮箱-要求集 1.8\""
description: ''
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 9bf71808b538f3048b6cc5f74a1bb0b9f88bea68
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163631"
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
| [过程](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#diagnostics) | ReadItem | 撰写<br>读取 | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#ewsurl) | ReadItem | 撰写<br>读取 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>读取 | [项](/javascript/api/outlook/office.item?view=outlook-js-1.8) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#mastercategories) | ReadWriteMailbox | 撰写<br>读取 | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#resturl) | ReadItem | 撰写<br>读取 | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#userprofile) | ReadItem | 撰写<br>读取 | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.8) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| 方法 | 最低<br>权限级别 | 型号 | 最低<br>要求集 |
|---|---|---|:---:|
| [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | 撰写<br>读取 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#converttoewsid-itemid--restversion-) | 受限 | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#converttolocalclienttime-timevalue-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#converttorestid-itemid--restversion-) | 受限 | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#converttoutcclienttime-input-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#displayappointmentform-itemid-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#displaymessageform-itemid-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#displaynewappointmentform-parameters-) | ReadItem | 读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [Office.context.mailbox.displaynewmessageform](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#displaynewmessageform-parameters-) | ReadItem | 撰写<br>读取 | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#getcallbacktokenasync-options--callback-) | ReadItem | 撰写<br>读取 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#getcallbacktokenasync-callback--usercontext-) | ReadItem | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#removehandlerasync-eventtype--options--callback-) | ReadItem | 撰写<br>读取 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>活动

您可以分别使用[addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#addhandlerasync-eventtype--handler--options--callback-)和[removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#removehandlerasync-eventtype--options--callback-)订阅和取消订阅以下事件。

| 事件 | 说明 | 最低<br>要求集 |
|---|---|:---:|
|`ItemChanged`| 在任务窗格固定时，将选择不同的 Outlook 项进行查看。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
