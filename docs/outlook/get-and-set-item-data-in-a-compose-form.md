---
title: 在 Outlook 的撰写窗体中获取和设置项目数据
description: 在撰写应用场景中获取或设置 Outlook 加载项中项的不同属性，包括收件人、主题、正文和约会地点和时间。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 3b82f418ffa2820e5f8cf04805a62b0d85691420
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325445"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a>在 Outlook 的撰写窗体中获取和设置项目数据

了解如何在撰写方案中获取或设置 Outlook 外接程序中项目的不同属性，包括收件人、主题、正文和约会地点和时间。

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a>获取和设置撰写加载项的项目属性

在撰写窗体中，您可以如同在阅读窗体中一样，获取在同一类型的项目上公开的大部分属性（如参与者、收件人、主题和正文），还可以获取仅与撰写窗体（而非阅读窗体）相关的一些其他属性（正文、密件抄送）。

对于大多数属性，由于 Outlook 外接程序和用户可能会同时修改用户界面中的同一个属性，获取和设置属性的方法将为异步。表 1 列出了项目级别属性以及用于在撰写窗体中获取和设置属性的相应异步方法。[item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [item.conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性是例外，因为用户无法修改。您可以使用与在阅读窗体中相同的编程方式，在撰写窗体中直接从父对象获取这些属性。

除了在 Office JavaScript API 中访问项目属性之外，还可以使用 Exchange Web 服务（EWS）访问项目级属性。 通过 **ReadWriteMailbox** 权限，可以使用 [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法访问 EWS 操作 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 和 [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)，以获取和设置用户邮箱中的一个或多个项目的更多属性。

`makeEwsRequestAsync` 函数在撰写窗体和阅读窗体中均可用。 有关 **ReadWriteMailbox** 权限以及通过 Office 加载项平台访问 EWS 的详细信息，请参阅[了解 Outlook 加载项权限](understanding-outlook-add-in-permissions.md)和[从 Outlook 加载项中调用 Web 服务](web-services.md)。

**表 1. 在撰写窗体中获取或设置项目属性的异步方法**

<br/>

| 属性 | 属性类型 | 获取的异步方法 | 设置的异步方法 |
|:-----|:-----|:-----|:-----|
|[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[收件人](/javascript/api/outlook/office.Recipients)|[Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)|[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)|
|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Body](/javascript/api/outlook/office.Body)|[Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)|[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-), [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)|
|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[时间](/javascript/api/outlook/office.Time)|[Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-)|[Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)|
|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[位置](/javascript/api/outlook/office.Location)|[Location.getAsync](/javascript/api/outlook/office.Location#getasync-options--callback-)|[Location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-)|
|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|时间|Time.getAsync|Time.setAsync|
|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Subject](/javascript/api/outlook/office.Subject)|[Subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-)|[Subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)|
|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|

## <a name="see-also"></a>另请参阅

- [创建适用于撰写窗体的 Outlook 加载项](compose-scenario.md)
- [了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)
- [从 Outlook 外接程序调用 Web 服务](web-services.md)
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](item-data.md)
