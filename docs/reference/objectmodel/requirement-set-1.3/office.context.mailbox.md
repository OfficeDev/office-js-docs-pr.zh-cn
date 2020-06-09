---
title: "\"Context.subname\"-\"邮箱-要求集 1.3\""
description: Outlook 邮箱 API 要求集的邮箱对象模型的版本为1.3。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: bbf97162b3ba2e6ea7694d2ce4de2dd4e09d580f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610473"
---
# <a name="mailbox-requirement-set-13"></a>邮箱（要求集1.3）

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
| [过程](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#diagnostics) | ReadItem | 撰写<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.3) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#ewsurl) | ReadItem | 撰写<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>Read | [项](/javascript/api/outlook/office.item?view=outlook-js-1.3) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#userprofile) | ReadItem | 撰写<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| 方法 | 最低<br>权限级别 | 型号 | 最低<br>要求集 |
|---|---|---|:---:|
| [Office.context.mailbox.converttoewsid （itemId，Office.mailboxenums.restversion）](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#converttoewsid-itemid--restversion-) | 受限 | 撰写<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime （timeValue）](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#converttolocalclienttime-timevalue-) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [Office.context.mailbox.converttorestid （itemId，Office.mailboxenums.restversion）](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#converttorestid-itemid--restversion-) | 受限 | 撰写<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime （输入）](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#converttoutcclienttime-input-) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#displayappointmentform-itemid-) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#displaymessageform-itemid-) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#displaynewappointmentform-parameters-) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#getcallbacktokenasync-callback--usercontext-) | ReadItem | 撰写<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
