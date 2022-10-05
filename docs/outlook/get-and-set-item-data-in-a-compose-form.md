---
title: 在 Outlook 的撰写窗体中获取和设置项目数据
description: 在撰写应用场景中获取或设置 Outlook 加载项中项的不同属性，包括收件人、主题、正文和约会地点和时间。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2ae4b6a30d08199207faf89079c57fbff46d6a0e
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467236"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>在 Outlook 的撰写窗体中获取和设置项目数据

了解如何在撰写方案中获取或设置 Outlook 外接程序中项目的不同属性，包括收件人、主题、正文和约会地点和时间。

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>获取和设置撰写加载项的项目属性

在撰写窗体中，您可以如同在阅读窗体中一样，获取在同一类型的项目上公开的大部分属性（如参与者、收件人、主题和正文），还可以获取仅与撰写窗体（而非阅读窗体）相关的一些其他属性（正文、密件抄送）。

For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The  [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [item.conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.

除了访问 Office JavaScript API 中的项属性之外，还可以使用 Exchange Web Services (EWS) 访问项级属性。 借助 **读/写入邮箱** 权限，可以使用 [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法访问 EWS 操作、 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 和 [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)，以获取和设置用户邮箱中项或项的更多属性。

该 `makeEwsRequestAsync` 方法在撰写和读取窗体中均可用。 有关 **读/写邮箱** 权限以及通过 Office 加载项平台访问 EWS 的详细信息，请参阅 Outlook 外接程序中的 [“了解 Outlook 外接程序”](understanding-outlook-add-in-permissions.md) 权限和 [调用 Web 服务](web-services.md)。

**表 1. 在撰写窗体中获取或设置项目属性的异步方法**

| 属性 | 属性类型 | 获取的异步方法 | 要设置的异步方法 |
|:-----|:-----|:-----|:-----|
|[bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[收件人](/javascript/api/outlook/office.recipients)|[Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))|[Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1)), [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))|
|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Body](/javascript/api/outlook/office.body)|[Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))|[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)), [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))|
|[cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Time](/javascript/api/outlook/office.time)|[Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1))|[Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))|
|[location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[位置](/javascript/api/outlook/office.location)|[Location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1))|[Location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))|
|[optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|时间|Time.getAsync|Time.setAsync|
|[subject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Subject](/javascript/api/outlook/office.subject)|[Subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1))|[Subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))|
|[to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>另请参阅

- [创建适用于撰写窗体的 Outlook 加载项](compose-scenario.md)
- [了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)
- [从 Outlook 外接程序调用 Web 服务](web-services.md)
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](item-data.md)
