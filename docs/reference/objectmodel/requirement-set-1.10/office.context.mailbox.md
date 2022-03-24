---
title: Office.context.mailbox - 要求集 1.10
description: Outlook邮箱 API 要求集 1.10 版本的邮箱对象模型。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 12806868e150471846c477166a478d06b803ec9a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746095"
---
# <a name="mailbox-requirement-set-110"></a>mailbox (requirement set 1.10) 

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-diagnostics-member) | ReadItem | 撰写<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-ewsurl-member) | ReadItem | 撰写<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>Read | [项目](/javascript/api/outlook/office.item?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-mastercategories-member) | ReadWriteMailbox | 撰写<br>Read | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.10&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-resturl-member) | ReadItem | 撰写<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-userprofile-member) | ReadItem | 撰写<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>方法

| 方法 | 最小值<br>权限级别 | 模式 | 最小值<br>要求集 |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) | ReadItem | 撰写<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId，restVersion) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-converttoewsid-member(1)) | 受限 | 撰写<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-converttolocalclienttime-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId， restVersion) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-converttorestid-member(1)) | 受限 | 撰写<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (输入) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-converttoutcclienttime-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displayappointmentform-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync (itemId， [options]， [callback]) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displayappointmentformasync-member(1)) | ReadItem | 撰写<br>Read | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaymessageform-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync (itemId， [options]， [callback]) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaymessageformasync-member(1)) | ReadItem | 撰写<br>Read | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaynewappointmentform-member(1)) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync (参数，[options]， [callback]) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaynewappointmentformasync-member(1)) | ReadItem | Read | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewMessageForm (参数) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaynewmessageform-member(1)) | ReadItem | Read | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync (参数，[options]， [callback]) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaynewmessageformasync-member(1)) | ReadItem | Read | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(1)) | ReadItem | 撰写<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(2)) | ReadItem | 撰写<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-getuseridentitytokenasync-member(1)) | ReadItem | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-makeewsrequestasync-member(1)) | ReadWriteMailbox | 撰写<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) | ReadItem | 撰写<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>事件

分别使用 [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) 和 [removeHandlerAsync 订阅和取消订阅以下](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) 事件。

> [!IMPORTANT]
> 事件仅适用于任务窗格实现。

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.10&preserve-view=true) | 说明 | 最小值<br>要求集 |
|---|---|:---:|
|`ItemChanged`| 在任务窗格固定时，将选择不同的 Outlook 项进行查看。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
