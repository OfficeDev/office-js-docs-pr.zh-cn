---
title: Outlook 外接程序 API 要求集 1.2
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: a77c9c0399d42029b86595f119c51071dc3bf4ac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433542"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Outlook 外接程序 API 要求集 1.2

适用于 Office 的 JavaScript API 的 Outlook 加载项 API 子集包括可以在 Outlook 加载项中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。 

## <a name="whats-new-in-12"></a>1.2 中的新增功能有哪些？

要求集 1.2 包括[要求集 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) 的所有功能。它添加了外接程序在用户的游标中插入文本的功能，无论是在邮件的主题还是正文中皆可插入文本。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string)：以异步方式返回邮件主题或正文中的选定数据。
- 添加了 [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback)：以异步方式将数据插入到邮件的正文或主题中。
- 修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata)：将 `attachments` 属性添加到 `formData` 参数。
- 修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata)：将 `attachments` 属性添加到 `formData` 参数。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)