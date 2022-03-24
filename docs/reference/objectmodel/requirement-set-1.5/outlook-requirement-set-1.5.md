---
title: Outlook 加载项 API 要求集 1.5
description: 为邮箱 API 1.5 Outlook外接程序和 Office JavaScript API 引入的功能和 API。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: ae549d001b39b43a9b2f258f9282e6b0093f94b3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746599"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Outlook 外接程序 API 要求集 1.5

Outlook JavaScript API 的 Office 加载项 API 子集包括可在加载项中Outlook的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-15"></a>1.5 中的新增功能有哪些？

要求集 1.5 包括要求集 [1.4 的所有功能](../requirement-set-1.4/outlook-requirement-set-1.4.md)。 它还添加了下列功能。

- 添加了对[可固定任务窗格](../../../outlook/pinnable-taskpane.md)的支持。
- 添加了对 [REST API](../../../outlook/use-rest-api.md) 调用的支持。
- 添加了将附件标记为内联的功能。
- 添加了关闭任务窗格或对话框的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods)：添加支持事件的事件处理程序。
- 添加了 [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods)：删除支持的事件类型的事件处理程序。
- 添加了 [Office.EventType](office.md#eventtype-string)：指定与事件处理程序相关联的事件，并包括对 ItemChanged 事件的支持。
- 添加了 [Office.context.mailbox.restUrl](office.context.mailbox.md#properties)：获取此电子邮件帐户的 REST 终结点的 URL。
- 修改了 [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods)：添加了此方法的新版本（具有新签名） (`getCallbackTokenAsync([options], callback)`)。原始版本仍可用且未更改。
- 添加了 [Office.context.ui.closeContainer](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true#office-office-ui-closecontainer-member(1))。
- 修改了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods)：`options` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。
- 修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods)：`formData.attachments` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。
- 修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods)：`formData.attachments` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
