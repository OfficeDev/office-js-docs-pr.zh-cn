---
title: Outlook 加载项 API 要求集 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: f56cf4e13bdf3518ef14da6eca83b51abe82e50c
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597045"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Outlook 外接程序 API 要求集 1.5

Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-15"></a>1.5 中的新增功能有哪些？

要求集 1.5 包括[要求集 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) 的所有功能。它添加了下列功能。

- 添加了对[可固定任务窗格](../../../outlook/pinnable-taskpane.md)的支持。
- 添加了对 [REST API](../../../outlook/use-rest-api.md) 调用的支持。
- 添加了将附件标记为内联的功能。
- 添加了关闭任务窗格或对话框的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods)：添加支持事件的事件处理程序。
- 添加了[removeHandlerAsync](office.context.mailbox.md#methods)：删除受支持的事件类型的事件处理程序。
- 添加了 [Office.EventType](office.md#eventtype-string)：指定与事件处理程序相关联的事件，并包括对 ItemChanged 事件的支持。
- 添加了 [Office.context.mailbox.restUrl](office.context.mailbox.md#properties)：获取此电子邮件帐户的 REST 终结点的 URL。
- 修改了 [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods)：添加了此方法的新版本（具有新签名） (`getCallbackTokenAsync([options], callback)`)。原始版本仍可用且未更改。
- 添加了 [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--)。
- 修改了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods)：`options` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。
- 修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods)：`formData.attachments` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。
- 修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods)：`formData.attachments` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
