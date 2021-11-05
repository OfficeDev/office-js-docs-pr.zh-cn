---
title: Office.context.mailbox - 要求集 1.11
description: Outlook邮箱 API 要求集 1.11 版本的邮箱对象模型。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2932376bd5e31348cde4480af62d86edcaf1a2c3
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681781"
---
# <a name="mailbox-requirement-set-111"></a>邮箱 (要求集 1.11) 

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#diagnostics) | ReadItem | 撰写<br>读取 | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#ewsUrl) | ReadItem | 撰写<br>读取 | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [项](office.context.mailbox.item.md) | 受限 | 撰写<br>读取 | [项目](/javascript/api/outlook/office.item?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#masterCategories) | ReadWriteMailbox | 撰写<br>读取 | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#restUrl) | ReadItem | 撰写<br>读取 | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#userProfile) | ReadItem | 撰写<br>读取 | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Methods

| 方法 | 最小值<br>权限级别 | 模式 | 最小值<br>要求集 |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | ReadItem | 撰写<br>读取 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId， restVersion) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#convertToEwsId_itemId__restVersion_) | 受限 | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#convertToLocalClientTime_timeValue_) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId， restVersion) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#convertToRestId_itemId__restVersion_) | 受限 | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (输入) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#convertToUtcClientTime_input_) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#displayAppointmentForm_itemId_) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync (itemId， [options]， [callback]) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#displayAppointmentFormAsync_itemId__options__callback_) | ReadItem | 撰写<br>读取 | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#displayMessageForm_itemId_) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync (itemId， [options]， [callback]) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#displayMessageFormAsync_itemId__options__callback_) | ReadItem | 撰写<br>读取 | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#displayNewAppointmentForm_parameters_) | ReadItem | 读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync (参数，[options]， [callback]) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#displayNewAppointmentFormAsync_parameters__options__callback_) | ReadItem | 读取 | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewMessageForm (参数) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#displayNewMessageForm_parameters_) | ReadItem | 读取 | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync (参数，[options]， [callback]) ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#displayNewMessageFormAsync_parameters__options__callback_) | ReadItem | 读取 | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#getCallbackTokenAsync_options__callback_) | ReadItem | 撰写<br>读取 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#getCallbackTokenAsync_callback__userContext_) | ReadItem | 撰写<br>读取 | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#getUserIdentityTokenAsync_callback__userContext_) | ReadItem | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#makeEwsRequestAsync_data__callback__userContext_) | ReadWriteMailbox | 撰写<br>读取 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | ReadItem | 撰写<br>读取 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>活动

分别使用 [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) 和 [removeHandlerAsync 订阅和取消订阅以下](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#removeHandlerAsync_eventType__options__callback_) 事件。

> [!IMPORTANT]
> 事件仅适用于任务窗格实现。

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.11&preserve-view=true) | 说明 | 最小值<br>要求集 |
|---|---|:---:|
|`ItemChanged`| 在任务窗格固定时，将选择不同的 Outlook 项进行查看。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |