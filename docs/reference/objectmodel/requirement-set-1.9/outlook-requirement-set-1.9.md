---
title: Outlook 加载项 API 要求集1。9
description: 适用于 Outlook 外接程序 API 的要求集1.9。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: b2174052a60580a895ef82a4b5d8f00ed6899feb
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628053"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook 加载项 API 要求集1。9

Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

## <a name="whats-new-in-19"></a>1.9 中的新增功能有哪些？

要求集1.9 包括 [要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md)的所有功能。 它还添加了下列功能。

- 添加了新的用于追加发送、自定义属性和显示表单功能的 Api。
- 添加了对的支持 `Dialog.messageChild` 。

### <a name="change-log"></a>更改日志

- 添加了 [CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getall--)：将新函数添加到 `CustomProperties` 获取所有自定义属性的对象。
- 添加了 [messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)：添加了一个新方法，该方法将来自主机页（如任务窗格或无 UI 的函数文件）的邮件传递到从页面打开的对话框。
- 添加了 [ExtendedPermissions 清单元素](../../manifest/extendedpermissions.md)：将子元素添加到 [VersionOverrides](../../manifest/versionoverrides.md) 清单元素中。 对于加载项以支持 [附加发送功能](../../../outlook/append-on-send.md)，扩展权限 `AppendOnSend` 必须包含在扩展的权限集合中。
- 添加了 [displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentformasync-itemid--options--callback-)：向 `Mailbox` 显示现有约会的对象添加新函数。 这是方法的异步版本 `displayAppointmentForm` 。
- 添加了 [displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageformasync-itemid--options--callback-)：向 `Mailbox` 显示现有邮件的对象添加新函数。 这是方法的异步版本 `displayMessageForm` 。
- 添加了 [displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)：向 `Mailbox` 显示新约会窗体的对象添加新函数。 这是方法的异步版本 `displayNewAppointmentForm` 。
- 添加了 [displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageformasync-parameters--options--callback-)：向 `Mailbox` 显示新邮件窗体的对象添加新函数。 这是方法的异步版本 `displayNewMessageForm` 。
- 添加了 [appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-)：将新函数添加到 `Body` 在撰写模式下将数据追加到项目正文末尾的对象。
- 添加了 [displayReplyAllFormAsync](office.context.mailbox.item.md#methods)：将新函数添加到 `Item` 在阅读模式下显示 "全部答复" 窗体的对象。 这是方法的异步版本 `displayReplyAllForm` 。
- 添加了 [displayReplyFormAsync](office.context.mailbox.item.md#methods)：向 `Item` 在阅读模式下显示 "答复" 窗体的对象添加新函数。 这是方法的异步版本 `displayReplyForm` 。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
