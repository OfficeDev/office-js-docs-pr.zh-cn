---
title: Outlook 外接程序 API 要求集 1.6
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 22702448b82a108c401f9f81d3b8a321e14ead63
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814659"
---
# <a name="outlook-add-in-api-requirement-set-16"></a>Outlook 外接程序 API 要求集 1.6

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。

## <a name="whats-new-in-16"></a>1.6 中的新增功能有哪些？

要求集 1.6 包括[要求集 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) 的所有功能。 它还添加了下列功能。

- 为上下文外接程序添加了新 API，以获取用户选择用于激活外接程序的实体或 RegEx 匹配项。
- 添加了新 API，用于打开新邮件窗体。
- 添加了通过外接程序来确定用户邮箱的帐户类型的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods)：添加了一个新函数，该函数可用于获取用户选择的突出显示匹配项中的实体。 突出显示的匹配项适用于上下文外接程序。
- 添加了 [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods)：添加了一个新函数，该函数可用于返回突出显示匹配项中与清单 XML 文件中定义的正则表达式匹配的字符串值。 突出显示的匹配项适用于上下文外接程序。
- 添加了 [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods)：添加了一个新函数，该函数将打开新邮件窗体。
- 添加了 [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#properties)：向指示用户帐户类型的用户配置文件添加了一个新成员。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](/outlook/add-ins/quick-start)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
