---
title: Outlook 外接程序 API 要求集 1.2
description: 为邮箱 API 1.2 Outlook外接程序和 Office JavaScript API 引入的功能和 API。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: d643f0fdf07c5f22d8d863075b894cfc05b21363
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590398"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Outlook 外接程序 API 要求集 1.2

Outlook JavaScript API 的 Office 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-12"></a>1.2 中的新增功能有哪些？

要求集 1.2 包括要求集 [1.1 的所有功能](../requirement-set-1.1/outlook-requirement-set-1.1.md)。 它添加了外接程序在用户的游标中插入文本的功能，无论是在邮件的主题还是正文中皆可插入文本。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods)：以异步方式返回邮件主题或正文中的选定数据。
- 添加了 [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods)：以异步方式将数据插入到邮件的正文或主题中。
- 修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods)：将 `attachments` 属性添加到 `formData` 参数。
- 修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods)：将 `attachments` 属性添加到 `formData` 参数。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
