---
title: "\"Context\"-\"邮箱-预览要求集\""
description: Outlook 邮箱 API preview 要求集邮箱对象模型的版本。
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 2aafc817555a59dcdadac2cc34bb69e487439639
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778667"
---
# <a name="mailbox-preview-requirement-set"></a>邮箱（预览要求集）

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
| [过程](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#diagnostics) | ReadItem | 撰写<br>阅读 | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#ewsurl) | ReadItem | 撰写<br>阅读 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>阅读 | [项](/javascript/api/outlook/office.item?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#mastercategories) | ReadWriteMailbox | 撰写<br>阅读 | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-preview) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#resturl) | ReadItem | 撰写<br>阅读 | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#userprofile) | ReadItem | 撰写<br>阅读 | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| 方法 | 最低<br>权限级别 | 型号 | 最低<br>要求集 |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | 撰写<br>阅读 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [Office.context.mailbox.converttoewsid （itemId，Office.mailboxenums.restversion）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttoewsid-itemid--restversion-) | 受限 | 撰写<br>阅读 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime （timeValue）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttolocalclienttime-timevalue-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [Office.context.mailbox.converttorestid （itemId，Office.mailboxenums.restversion）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttorestid-itemid--restversion-) | 受限 | 撰写<br>阅读 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime （输入）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttoutcclienttime-input-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentform-itemid-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync （itemId，[options]，[callback]）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentform-itemid--options--callback-) | ReadItem | 撰写<br>阅读 | [预览](outlook-requirement-set-preview.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageform-itemid-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync （itemId，[options]，[callback]）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageform-itemid--options--callback-) | ReadItem | 撰写<br>阅读 | [预览](outlook-requirement-set-preview.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentform-parameters-) | ReadItem | 阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync （parameters、[options]、[callback]）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentform-parameters--options--callback-) | ReadItem | 阅读 | [预览](outlook-requirement-set-preview.md) |
| [Office.context.mailbox.displaynewmessageform （参数）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageform-parameters-) | ReadItem | 撰写<br>阅读 | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync （parameters、[options]、[callback]）](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageform-parameters--options--callback-) | ReadItem | 撰写<br>阅读 | [预览](outlook-requirement-set-preview.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getcallbacktokenasync-options--callback-) | ReadItem | 撰写<br>阅读 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getcallbacktokenasync-callback--usercontext-) | ReadItem | 撰写<br>阅读 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 撰写<br>阅读 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | ReadItem | 撰写<br>阅读 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>活动

您可以分别使用[addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-)和[removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-)订阅和取消订阅以下事件。

| 事件 | 说明 | 最低<br>要求集 |
|---|---|:---:|
|`ItemChanged`| 在任务窗格固定时，将选择不同的 Outlook 项进行查看。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
|`OfficeThemeChanged`| 邮箱上的 Office 主题已更改。 | [预览](../preview-requirement-set/outlook-requirement-set-preview.md) |
