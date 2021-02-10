---
title: Outlook 外接程序 API 预览要求集
description: Outlook 外接程序当前处于预览阶段的功能和 API。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 92ba3510af0c8b9ebdf9ca4368c889b821a9cb3b
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173953"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

Office JavaScript API 的 Outlook 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!IMPORTANT]
> 本文档适用于 **预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> 你可能能够通过在 [Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)租户上配置定向版本来预览 Outlook 网页版中的功能。 "配置预览访问"在此页面上针对适用的功能进行说明。
>
> 对于其他功能，你可能能够通过完成和提交此表单，请求访问 Outlook 网页版预览位，使用 Microsoft 365 [帐户](https://aka.ms/OWAPreview)。 这些功能上会指出"请求预览访问"。

预览要求集包括要求集 [1.9 的所有功能](../requirement-set-1.9/outlook-requirement-set-1.9.md)。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>对受信息权限管理中心 IRM 保护的项目 (加载项) 

加载项现在可以在受 IRM 保护的项目上激活。 若要启用此功能，租户管理员需要通过设置 Office 中的"允许编程访问自定义策略"选项来 `OBJMODEL` 启用使用权限。  有关详细信息 [，请参阅使用](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 权限和说明。

**适用于：Windows** 版 Outlook，从内部版本 13229.10000 (连接到 Microsoft 365 订阅) 

<br>

---

---

### <a name="additional-calendar-properties"></a>其他日历属性

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

新增了一个对象，该对象代表撰写模式下约会的全天事件属性。

**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) 

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

新增了一个对象，该对象代表撰写模式下约会的敏感度。

**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) 

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

添加了一个新属性，表示约会是全天事件。

**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) 

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

添加了一个新属性，表示约会的敏感度。

**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) 

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

添加了一个新枚举 `AppointmentSensitivityType` ，该枚举代表约会可用的敏感度选项。

**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) 

<br>

---

---

### <a name="event-based-activation"></a>基于事件的激活

增加了对 Outlook 外接程序中基于事件的激活功能的支持。请参阅 [配置 Outlook 外接程序进行基于事件的激活](../../../outlook/autolaunch.md) 以了解更多信息。

#### <a name="launchevent-extension-point"></a>[LaunchEvent 扩展点](../../manifest/extensionpoint.md#launchevent-preview)

向 `LaunchEvent` 清单添加了扩展点支持。 它配置基于事件的激活功能。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="launchevents-manifest-element"></a>[LaunchEvents 清单元素](../../manifest/launchevents.md)

向 `LaunchEvents` 清单添加了元素。 它支持配置基于事件的激活功能。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="runtimes-manifest-element"></a>[运行时清单元素](../../manifest/runtimes.md)

向清单元素添加了 Outlook `Runtimes` 支持。 它引用基于事件的激活功能所需的 HTML 和 JavaScript 文件。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>与可操作邮件集成

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) 

<br>

---

---

### <a name="mail-signature"></a>邮件签名

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

向对象添加了一个新函数，该函数在撰写模式下添加或替换 `Body` 项目正文中的签名。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods)

添加了一个新函数，该函数在撰写模式下禁用发送邮箱的客户端签名。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

添加了一个新函数，该函数获取撰写模式下邮件的撰写类型。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

新增了一个函数，该函数检查在撰写模式下是否对项目启用了客户端签名。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officemailboxenumscomposetype"></a>[Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

添加了在撰写模式下 `ComposeType` 可用的新枚举。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

<br>

---

---

### <a name="notification-messages-with-actions"></a>包含操作的通知邮件

此功能允许外接程序在默认消除操作之外包含包含自定义操作 **的通知** 消息。 在新式 Outlook 网页中，此功能仅在撰写模式下可用。

#### <a name="officenotificationmessagedetailsactions"></a>[Office.NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

添加了一个新属性，允许您使用自定义操作 `InsightMessage` 添加通知。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) 

#### <a name="officenotificationmessageaction"></a>[Office.NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

添加了一个新对象，用于定义通知的自定义 `InsightMessage` 操作。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) 

#### <a name="officemailboxenumsactiontype"></a>[Office.MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype)

添加了一个新枚举 `ActionType` 。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) 

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[Office.MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

向枚举添加了 `InsightMessage` 一个新 `ItemNotificationMessageType` 类型。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) 

<br>

---

---

### <a name="office-theme"></a>Office 主题

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

增加了获取 Office 主题的功能。

**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) 

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。

**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) 

<br>

---

---

### <a name="session-data"></a>会话数据

#### <a name="officesessiondata"></a>[Office.SessionData](/javascript/api/outlook/office.sessiondata)

添加了一个新对象，该对象代表项目的会话数据。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) 

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

添加了一个新属性，用于管理撰写模式下项目的会话数据。

**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) 

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
