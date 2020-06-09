---
title: Outlook 外接程序 API 要求集 1.3
description: 作为邮箱 API 1.3 的一部分引入的 Outlook 外接程序和 Office JavaScript Api 的功能和 Api。
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 2f51a275e00853b2b3626c710a4c072a83ba8c0a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611994"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Outlook 外接程序 API 要求集 1.3

Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-13"></a>1.3 中的新增功能有哪些？

要求集 1.3 包括[要求集 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) 的所有功能。它添加了下列功能。

- 添加了对[外接程序命令](../../../outlook/add-in-commands-for-outlook.md)的支持。
- 添加了保存或关闭正在撰写的项目的功能。
- 增强的[Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)对象，允许外接程序获取或设置整个正文。
- 添加了在 EWS 和 REST 格式之间转换 ID 的转换方法。
- 添加了将通知邮件添加到项目的信息栏中的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-)：使用指定格式返回当前正文。
- 添加了 [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#setasync-data--options--callback-)：将整个正文替换为指定文本。
- 添加了 [Event](/javascript/api/office/office.addincommands.event) 对象：作为参数传递到 Outlook 外接程序中的无用户界面命令函数。用来表示处理已完成。
- 添加了 [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods)：关闭正在撰写的当前项。
- 添加了 [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods)：异步保存项目。
- 添加了 [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties)：获取项目的通知邮件。
- 添加了 [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods)：将项目 ID 格式化（从 REST 转换为 EWS 格式）。
- 添加了 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods)：将项目 ID 格式化（从 EWS 转换为 REST 格式）。
- 添加了 [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3)：为约会或邮件指定通知邮件类型。
- 添加了 [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)：指定对应于 REST 格式的项目 ID 的 REST API 的版本。
- 添加了 [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3) 对象：提供用于访问 Outlook 外接程序中的通知邮件的方法。
- 添加了 [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3) 类型：由 `NotificationMessages.getAllAsync` 方法返回。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
