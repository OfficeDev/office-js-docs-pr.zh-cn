---
title: Outlook 外接程序 API 要求集 1.1
description: 为外接程序和 Outlook引入的功能和 API Office作为邮箱 API 1.1 的一部分。
ms.date: 12/17/2019
ms.localizationpriority: medium
ms.openlocfilehash: 74a6b2561cf7d0ec28d97810d7337fe909a30837
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746677"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Outlook 外接程序 API 要求集 1.1

Outlook JavaScript API 的 Office 加载项 API 子集包括可在加载项中Outlook的对象、方法、属性和事件。 Outlook JavaScript API 1.1 (Mailbox 1.1) 是首版 API。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-11"></a>1.1 中的新增功能有哪些？

要求集 1.1 包括所有通用 API 要求集，这些[通用 API](../../requirement-sets/office-add-in-requirement-sets.md) 要求集Outlook。 它添加了外接程序访问邮件和约会的正文以及修改当前项的功能。

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
- 添加了[Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) 行。
- 添加了 [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true)：指定约会收件人的类型。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
