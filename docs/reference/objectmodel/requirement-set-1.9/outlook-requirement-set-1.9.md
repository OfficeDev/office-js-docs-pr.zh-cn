---
title: Outlook外接程序 API 要求集 1.9
description: 加载项 API 要求集 1.9 Outlook 1.9。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 4448a7391e2d829fa95fa72392bf22867fafe7a7
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936793"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook外接程序 API 要求集 1.9

Outlook JavaScript API 的 Office API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。

## <a name="whats-new-in-19"></a>1.9 中的新增功能是什么？

要求集 1.9 包括要求集 [1.8 的所有功能](../requirement-set-1.8/outlook-requirement-set-1.8.md)。 它还添加了下列功能。

- 添加了用于附加 Onss、自定义属性和显示表单功能的新 API。
- 添加了对 `Dialog.messageChild` 的支持。

### <a name="change-log"></a>更改日志

- 添加了 [CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getAll__)：向获取所有自定义属性的对象添加了 `CustomProperties` 一个新函数。
- 添加了 [Dialog.messageChild：](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)添加了一个新方法，该方法将邮件从主机页（如任务窗格或无 UI 函数文件）发送到从该页面打开的对话框。
- 添加了 [ExtendedPermissions 清单元素](../../manifest/extendedpermissions.md)：向 [VersionOverrides](../../manifest/versionoverrides.md) 清单元素添加了子元素。 若要使外接程序支持 [附加 Onss 发送](../../../outlook/append-on-send.md)功能，扩展权限必须包含在扩展 `AppendOnSend` 权限集合中。
- 添加了[Office.context.mailbox.displayAppointmentFormAsync：](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayAppointmentFormAsync_itemId__options__callback_)向显示现有约会 `Mailbox` 的对象添加新函数。 这是 方法的异步 `displayAppointmentForm` 版本。
- 添加了[Office.context.mailbox.displayMessageFormAsync：](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayMessageFormAsync_itemId__options__callback_)向显示现有邮件 `Mailbox` 的对象添加新函数。 这是 方法的异步 `displayMessageForm` 版本。
- 添加了[Office.context.mailbox.displayNewAppointmentFormAsync：](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayNewAppointmentFormAsync_parameters__options__callback_)向显示新约会窗体的对象添加新 `Mailbox` 函数。 这是 方法的异步 `displayNewAppointmentForm` 版本。
- 添加了[Office.context.mailbox.displayNewMessageFormAsync：](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayNewMessageFormAsync_parameters__options__callback_)向显示新邮件表单的对象添加新 `Mailbox` 函数。 这是 方法的异步 `displayNewMessageForm` 版本。
- 添加了[Office.context.mailbox.item.body.appendOnSendAsync：](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendOnSendAsync_data__options__callback_)向在撰写模式下将数据追加到项目正文末尾 `Body` 的对象添加新函数。
- 添加了[Office.context.mailbox.item.displayReplyAllFormAsync：](office.context.mailbox.item.md#methods)向在阅读模式下显示"全部答复"窗体的对象添加新 `Item` 函数。 这是 方法的异步 `displayReplyAllForm` 版本。
- 添加了[Office.context.mailbox.item.displayReplyFormAsync：](office.context.mailbox.item.md#methods)向在阅读模式下显示"答复"窗体的对象添加新 `Item` 函数。 这是 方法的异步 `displayReplyForm` 版本。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
