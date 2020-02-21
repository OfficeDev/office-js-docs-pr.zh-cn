---
title: Outlook 外接程序 API 要求集 1.1
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 2ecd337cd838cd6dd9deb4fe5e77ee789106f3f9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165466"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook 外接程序 API 要求集 1.1

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。 Outlook JavaScript API 1.1 （邮箱1.1）是第一个 API 版本。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-11"></a>1.1 中的新增功能有哪些？

要求集1.1 包括在 Outlook 中支持的所有[通用 API 要求集](../../requirement-sets/office-add-in-requirement-sets.md)。 它添加了外接程序访问邮件和约会的正文以及修改当前项的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) 对象：提供用于在 Outlook 外接程序中添加和更新项目内容的方法。
- 添加了 [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的会议地点的方法。
- 添加了 [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的收件人的方法。
- 添加了 [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的主题的方法。
- 添加了 [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的会议开始或结束时间的方法。
- 添加了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods)：将文件作为附件添加到邮件或约会。
- 添加了 [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods)：将 Exchange 项目（如邮件）作为附件添加到邮件或约会。
- 添加了 [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods)：将附件从邮件或约会中删除。
- 添加了 [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties)：获取一个提供用于处理项目正文的方法的对象。
- 添加了邮件的["密件抄送"](office.context.mailbox.item.md#properties)行。
- 添加了 [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1)：指定约会收件人的类型。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
