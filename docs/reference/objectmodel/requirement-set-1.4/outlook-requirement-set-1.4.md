---
title: Outlook 加载项 API 要求集 1.4
description: 作为邮箱 API 1.4 Outlook外接程序和 Office JavaScript API 引入的功能和 API。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: b00413ef4c7f862a125c4a5a1d2190d4d60e87bf
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671252"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Outlook 外接程序 API 要求集 1.4

Outlook JavaScript API 的 Office 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-14"></a>1.4 中的新增功能有哪些？

要求集 1.4 包括要求集 [1.3 的所有功能](../requirement-set-1.3/outlook-requirement-set-1.3.md)。 它添加了对 `Office.ui` 命名空间的访问权限。

### <a name="change-log"></a>更改日志

- 添加了[Office.context.ui.displayDialogAsync：](/javascript/api/office/office.ui#displayDialogAsync_startAddress__options__callback_)在应用程序Office对话框。
- 添加了 [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageParent_message__messageOptions_)：将对话框中的消息传送到其父页/开始页。
- 添加了 [Dialog](/javascript/api/office/office.dialog) 对象：调用 [`displayDialogAsync`](/javascript/api/office/office.ui#displayDialogAsync_startAddress__options__callback_) 方法时返回的对象。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
