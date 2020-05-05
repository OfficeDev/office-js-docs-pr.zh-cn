---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序的预览中的功能和 Api。
ms.date: 05/04/2020
localization_priority: Normal
ms.openlocfilehash: 433abf5ea15cfe9ee7cefa579b1ad965fdeb93cf
ms.sourcegitcommit: 800dacf0399465318489c9d949e259b5cf0f81ca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/05/2020
ms.locfileid: "44022181"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!IMPORTANT]
> 本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

### <a name="additional-calendar-properties"></a>其他日历属性

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

在撰写模式下添加了一个代表约会全天事件属性的新对象。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

添加了一个新对象，该对象表示在撰写模式下约会的敏感度。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemisalldayevent"></a>[IsAllDayEvent 的 Office。](office.context.mailbox.item.md#properties)

添加了一个新属性，该属性表示约会是否为全天事件。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemsensitivity"></a>["Context"。项目敏感度](office.context.mailbox.item.md#properties)

添加了一个表示约会敏感度的新属性。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[MailboxEnums. AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

添加了一个代表`AppointmentSensitivityType`约会上可用的敏感度选项的新枚举。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

<br>

---

---

### <a name="append-on-send"></a>发送时追加

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[AppendOnSendAsync 的 "."](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

向`Body`对象添加了一个新函数，该函数在撰写模式下将数据追加到项正文的末尾。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="extendedpermissions"></a>[ExtendedPermissions](../../manifest/extendedpermissions.md)

向清单添加了一个新元素，其中`AppendOnSend`扩展权限必须包含在扩展权限的集合中。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

<br>

---

---

### <a name="integration-with-actionable-messages"></a>与可操作邮件集成

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）

<br>

---

---

### <a name="mail-signature"></a>邮件签名

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[SetSignatureAsync 的 "."](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

向`Body`对象添加了一个新函数，该函数在撰写模式下添加或替换项目正文中的签名。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[DisableClientSignatureAsync 的 Office。](office.context.mailbox.item.md#methods)

添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[GetComposeTypeAsync 的 Office。](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[IsClientSignatureEnabledAsync 的 Office。](office.context.mailbox.item.md#methods)

添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officemailboxenumscomposetype"></a>[MailboxEnums. ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

添加了一个新`ComposeType`枚举，该枚举在撰写模式中可用。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

<br>

---

---

### <a name="office-theme"></a>Office 主题

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

增加了获取 Office 主题的功能。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

<br>

---

---

### <a name="online-meeting-provider-integration"></a>联机会议提供程序集成

添加了对 Outlook 移动外接程序中的联机会议集成的支持。有关详细信息，请参阅为[联机会议提供商创建 Outlook mobile 外接程序](../../../outlook/online-meeting.md)以了解详细信息。

#### <a name="mobileonlinemeetingcommandsurface-extension-point"></a>[MobileOnlineMeetingCommandSurface 扩展点](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

向`MobileOnlineMeetingCommandSurface`清单添加了扩展点。 它定义联机会议集成。

**适用于**： Outlook on Android （连接到 Office 365 订阅）

<br>

---

---

### <a name="sso"></a>SSO

#### <a name="officeruntimeauthgetaccesstoken"></a>[OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
