---
title: Outlook 外接程序 API 要求集 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: cd284a5871139b7f6bf006a9deb3671a937682f6
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870070"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook 外接程序 API 要求集 1.1

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。 

## <a name="whats-new-in-11"></a>1.1 中的新增功能有哪些？

要求集 1.1 包括要求集 1.0 的所有功能。它添加了外接程序访问邮件和约会的正文以及修改当前项的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Body](/javascript/api/outlook_1_1/office.body) 对象：提供用于在 Outlook 外接程序中添加和更新项目内容的方法。
- 添加了 [Location](/javascript/api/outlook_1_1/office.location) 对象：提供用于获取和设置 Outlook 外接程序中的会议地点的方法。
- 添加了 [Recipients](/javascript/api/outlook_1_1/office.recipients) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的收件人的方法。
- 添加了 [Subject](/javascript/api/outlook_1_1/office.subject) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的主题的方法。
- 添加了 [Time](/javascript/api/outlook_1_1/office.time) 对象：提供用于获取和设置 Outlook 外接程序中的会议开始或结束时间的方法。
- 添加了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback)：将文件作为附件添加到邮件或约会。
- 添加了 [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback)：将 Exchange 项目（如邮件）作为附件添加到邮件或约会。
- 添加了 [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback)：将附件从邮件或约会中删除。
- 添加了 [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body)：获取一个提供用于处理项目正文的方法的对象。
- 添加了邮件的["密件抄送"](office.context.mailbox.item.md#bcc-recipients)行。
- 添加了 [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype)：指定约会收件人的类型。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](/outlook/add-ins/quick-start)
