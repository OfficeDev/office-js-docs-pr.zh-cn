---
title: Office.context.mailbox - 要求集 1.7
description: Outlook邮箱 API 要求集 1.7 版本的邮箱对象模型。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 22ef0c2692cfb4abddc67a6adc26be99e099fa85
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590986"
---
# <a name="mailbox-requirement-set-17"></a>邮箱 (要求集 1.7) 

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#diagnostics) | ReadItem | 撰写<br>阅读 | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#ewsurl) | ReadItem | 撰写<br>阅读 | 字符串 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>阅读 | [项目](/javascript/api/outlook/office.item?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#resturl) | ReadItem | 撰写<br>阅读 | 字符串 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#userprofile) | ReadItem | 撰写<br>阅读 | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.7&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>方法

| 方法 | 最小值<br>权限级别 | 模式 | 最小值<br>要求集 |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | 撰写<br>阅读 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId， restVersion) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#converttoewsid-itemid--restversion-) | 受限 | 撰写<br>阅读 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId， restVersion) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#converttorestid-itemid--restversion-) | 受限 | 撰写<br>阅读 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (输入) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#converttoutcclienttime-input-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#displayappointmentform-itemid-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#displaymessageform-itemid-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | 阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm (参数) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#displaynewmessageform-parameters-) | ReadItem | 阅读 | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | 撰写<br>阅读 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | 撰写<br>阅读 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | 撰写<br>阅读 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>活动

可以分别使用 [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) 和 [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true#removehandlerasync-eventtype--options--callback-) 订阅和取消订阅以下事件。

> [!IMPORTANT]
> 事件仅适用于任务窗格实现。

| 事件 | 说明 | 最小值<br>要求集 |
|---|---|:---:|
|`ItemChanged`| 在任务窗格固定时，将选择不同的 Outlook 项进行查看。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
