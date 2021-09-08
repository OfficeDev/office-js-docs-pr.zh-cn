---
title: Outlook 外接程序 API 要求集 1.1
description: 为邮箱 API 1.1 Outlook外接程序和 Office JavaScript API 引入的功能和 API。
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: f93b6d582043641903b362121c6e5eaf89c2ad1c
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937481"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook 外接程序 API 要求集 1.1

Outlook JavaScript API 的 Office 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。 OutlookJavaScript API 1.1 (Mailbox 1.1) 是首版 API。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-11"></a>1.1 中的新增功能有哪些？

要求集 1.1 包括所有[通用 API](../../requirement-sets/office-add-in-requirement-sets.md)要求集，这些通用 API 要求集Outlook。 它添加了外接程序访问邮件和约会的正文以及修改当前项的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) 对象：提供用于在 Outlook 外接程序中添加和更新项目内容的方法。
- 添加了 [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true) 对象：提供用于获取和设置 Outlook 外接程序中的会议地点的方法。
- 添加了 [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的收件人的方法。
- 添加了 [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的主题的方法。
- 添加了 [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true) 对象：提供用于获取和设置 Outlook 外接程序中的会议开始或结束时间的方法。
- 添加了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods)：将文件作为附件添加到邮件或约会。
- 添加了 [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods)：将 Exchange 项目（如邮件）作为附件添加到邮件或约会。
- 添加了 [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods)：将附件从邮件或约会中删除。
- 添加了 [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties)：获取一个提供用于处理项目正文的方法的对象。
- 添加了[Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties)行。
- 添加了 [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true)：指定约会收件人的类型。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
