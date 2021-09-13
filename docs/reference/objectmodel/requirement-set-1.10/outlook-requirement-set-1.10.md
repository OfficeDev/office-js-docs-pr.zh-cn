---
title: Outlook外接程序 API 要求集 1.10
description: 加载项 API 要求集 1.10 Outlook 1.10。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: b54d327d37acd7b2c7fcff100cc7dbe7a39187c0
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152618"
---
# <a name="outlook-add-in-api-requirement-set-110"></a>Outlook外接程序 API 要求集 1.10

Outlook JavaScript API 的 Office 加载项 API 子集包括可在加载项中Outlook的对象、方法、属性和事件。

## <a name="whats-new-in-110"></a>1.10 中的新增功能是什么？

要求集 1.10 包括要求集 [1.9 的所有功能](../requirement-set-1.9/outlook-requirement-set-1.9.md)。 它还添加了下列功能。

- 添加了用于基于事件的 [激活和邮件](../../../outlook/autolaunch.md) 签名功能的新 API。
- 添加了在通知邮件中添加自定义操作的能力。

### <a name="change-log"></a>更改日志

- 添加了 [LaunchEvent 扩展点](../../manifest/extensionpoint.md#launchevent)：添加了新的受支持的 ExtensionPoint 类型。 它配置基于事件的激活功能。
- 添加了 [LaunchEvents manifest 元素](../../manifest/launchevents.md)：添加了一个清单元素以支持配置基于事件的激活功能。
- 修改后的[运行时清单元素](../../manifest/runtimes.md)：添加Outlook支持。 它引用基于事件的激活功能所需的 HTML 和 JavaScript 文件。
- 添加了[Office.context.mailbox.item.body.setSignatureAsync：](/javascript/api/outlook/office.body?view=outlook-js-1.10&preserve-view=true#setSignatureAsync_data__options__callback_)向 对象添加新 `Body` 函数。 它在撰写模式下添加或替换项目正文中的签名。
- 添加了[Office.context.mailbox.item.disableClientSignatureAsync：](office.context.mailbox.item.md#methods)添加了一个新函数，该函数在撰写模式下禁用发送邮箱的客户端签名。
- 添加了[Office.context.mailbox.item.getComposeTypeAsync：](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getComposeTypeAsync_options__callback_)添加了一个新函数，该函数获取撰写模式下邮件的撰写类型。
- 添加了[Office.context.mailbox.item.isClientSignatureEnabledAsync：](office.context.mailbox.item.md#methods)添加了一个新函数，该函数检查在撰写模式下是否对项目启用了客户端签名。
- 添加了[Office。MailboxEnums.ActionType：](/javascript/api/outlook/office.mailboxenums.actiontype)添加新枚举。 它表示通知邮件中的自定义操作的类型。
- 添加了[Office.MailboxEnums.ComposeType：](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-1.10&preserve-view=true)添加了一个在撰写模式下可用的新枚举。
- 添加了[Office。MailboxEnums.ItemNotificationMessageType.InsightMessage：](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)向 `ItemNotificationMessageType` 枚举添加新类型。 它表示具有自定义操作的通知消息。
- 添加了[Office。NotificationMessageAction：](/javascript/api/outlook/office.notificationmessageaction)添加新对象，以便你可以为通知定义自定义 `InsightMessage` 操作。
- 添加了[Office。NotificationMessageDetails.actions：](/javascript/api/outlook/office.notificationmessagedetails#actions)添加一个新属性，允许您使用自定义操作 `InsightMessage` 添加通知。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
