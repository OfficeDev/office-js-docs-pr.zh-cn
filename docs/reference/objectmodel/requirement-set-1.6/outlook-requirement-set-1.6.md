---
title: Outlook 外接程序 API 要求集 1.6
description: 作为邮箱 API 1.6 Outlook外接程序和 Office JavaScript API 引入的功能和 API。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: a552c362e247da7b36d14a0c32f557440a324977
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671735"
---
# <a name="outlook-add-in-api-requirement-set-16"></a>Outlook 外接程序 API 要求集 1.6

Outlook JavaScript API 的 Office 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-16"></a>1.6 中的新增功能有哪些？

要求集 1.6 包括要求集 [1.5 的所有功能](../requirement-set-1.5/outlook-requirement-set-1.5.md)。 它还添加了下列功能。

- 为上下文外接程序添加了新 API，以获取用户选择用于激活外接程序的实体或 RegEx 匹配项。
- 添加了新 API，用于打开新邮件窗体。
- 添加了通过外接程序来确定用户邮箱的帐户类型的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods)：添加了一个新函数，该函数可用于获取用户选择的突出显示匹配项中的实体。 突出显示的匹配项适用于上下文外接程序。
- 添加了 [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods)：添加了一个新函数，该函数可用于返回突出显示匹配项中与清单 XML 文件中定义的正则表达式匹配的字符串值。 突出显示的匹配项适用于上下文外接程序。
- 添加了 [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods)：添加了一个新函数，该函数将打开新邮件窗体。
- 添加了 [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true#accountType)：向指示用户帐户类型的用户配置文件添加了一个新成员。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
