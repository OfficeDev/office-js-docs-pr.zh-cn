---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序的预览中的功能和 Api。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: d91105e0cfbb97dc1a239e40b1c81adc4e76988b
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626594"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!IMPORTANT]
> 本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> 您可以通过 [在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)来预览 Web 上 Outlook 中的功能。 此页面上的 "配置预览访问权限" 对适用的功能进行了说明。
>
> 对于其他功能，你可以通过填写和提交 [此表单](https://aka.ms/OWAPreview)，使用 Microsoft 365 帐户请求对网站上的 Outlook 的预览位的访问权限。 这些功能上记录了 "请求预览访问"。

预览要求集包括 [要求集 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md)的所有功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>对受信息权限管理 (IRM) 保护的项的外接程序激活

现在，外接程序可以在受 IRM 保护的项上激活。 若要启用此功能，租户管理员需要 `OBJMODEL` 通过在 Office 中设置 " **允许编程访问** " 自定义策略选项来启用使用权限。 有关详细信息，请参阅 [使用权限和说明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 。

**适用于**： Windows 中的 Outlook，从内部版本 13229.10000 (连接到 Microsoft 365 订阅) 

<br>

---

---

### <a name="additional-calendar-properties"></a>其他日历属性

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

在撰写模式下添加了一个代表约会全天事件属性的新对象。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

添加了一个新对象，该对象表示在撰写模式下约会的敏感度。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

#### <a name="officecontextmailboxitemisalldayevent"></a>[IsAllDayEvent 的 Office。](office.context.mailbox.item.md#properties)

添加了一个新属性，该属性表示约会是否为全天事件。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

#### <a name="officecontextmailboxitemsensitivity"></a>["Context"。项目敏感度](office.context.mailbox.item.md#properties)

添加了一个表示约会敏感度的新属性。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[MailboxEnums. AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

添加了一个 `AppointmentSensitivityType` 代表约会上可用的敏感度选项的新枚举。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

<br>

---

---

### <a name="event-based-activation"></a>基于事件的激活

添加了对 Outlook 外接程序中基于事件的激活功能的支持。若要了解详细信息，请参阅 [配置 Outlook 外接程序以进行基于事件的激活](../../../outlook/autolaunch.md) 。

#### <a name="launchevent-extension-point"></a>[LaunchEvent 扩展点](../../manifest/extensionpoint.md#launchevent-preview)

`LaunchEvent`向清单添加了扩展点支持。 它配置基于事件的激活功能。

**中的可用**： Outlook 网页版 (新式， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) 

#### <a name="launchevents-manifest-element"></a>[LaunchEvents 清单元素](../../manifest/launchevents.md)

`LaunchEvents`向清单添加了元素。 它支持配置基于事件的激活功能。

**中的可用**： Outlook 网页版 (新式， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) 

#### <a name="runtimes-manifest-element"></a>[运行时清单元素](../../manifest/runtimes.md)

向清单元素添加了 Outlook 支持 `Runtimes` 。 它引用了基于事件的激活功能所需的 HTML 和 JavaScript 文件。

**中的可用**： Outlook 网页版 (新式， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) 

<br>

---

---

### <a name="integration-with-actionable-messages"></a>与可操作邮件集成

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (经典) 

<br>

---

---

### <a name="mail-signature"></a>邮件签名

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[SetSignatureAsync 的 "."](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

向对象添加了一个新函数 `Body` ，该函数在撰写模式下添加或替换项目正文中的签名。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) 

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[DisableClientSignatureAsync 的 Office。](office.context.mailbox.item.md#methods)

添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) 

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[GetComposeTypeAsync 的 Office。](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) 

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[IsClientSignatureEnabledAsync 的 Office。](office.context.mailbox.item.md#methods)

添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) 

#### <a name="officemailboxenumscomposetype"></a>[MailboxEnums. ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

添加了一个新枚举，该枚举 `ComposeType` 在撰写模式中可用。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) 

<br>

---

---

### <a name="notification-messages-with-actions"></a>包含操作的通知邮件

通过此功能，您的外接程序可以在默认 **取消** 操作之外包含具有自定义操作的通知消息。

#### <a name="officenotificationmessagedetailsactions"></a>[NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails#actions)

添加了一个新属性，您可以 `InsightMessage` 使用自定义操作添加通知。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) 

#### <a name="officenotificationmessageaction"></a>[NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

添加了一个新对象，可在其中为通知定义自定义操作 `InsightMessage` 。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) 

#### <a name="officemailboxenumsactiontype"></a>[MailboxEnums](/javascript/api/outlook/office.mailboxenums.actiontype)

添加了新枚举 `ActionType` 。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) 

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[ItemNotificationMessageType InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

向枚举添加了新类型 `InsightMessage` `ItemNotificationMessageType` 。

**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) 

<br>

---

---

### <a name="office-theme"></a>Office 主题

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

增加了获取 Office 主题的功能。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

<br>

---

---

### <a name="session-data"></a>会话数据

#### <a name="officesessiondata"></a>[SessionData](/javascript/api/outlook/office.sessiondata)

添加了一个代表项目的会话数据的新对象。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

#### <a name="officecontextmailboxitemsessiondata"></a>[SessionData 的 Office。](office.context.mailbox.item.md#properties)

添加了一个新属性以在撰写模式下管理项目的会话数据。

**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) 

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
