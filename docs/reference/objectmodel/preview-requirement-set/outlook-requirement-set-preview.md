---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序和 Office JavaScript Api 的预览中的功能和 Api。
ms.date: 03/26/2020
localization_priority: Normal
ms.openlocfilehash: 55de284932a53d2226258a15c86ead4f05361c30
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978618"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!IMPORTANT]
> 本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

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

**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[DisableClientSignatureAsync 的 Office。](office.context.mailbox.item.md#methods)

添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。

**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[GetComposeTypeAsync 的 Office。](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[IsClientSignatureEnabledAsync 的 Office。](office.context.mailbox.item.md#methods)

添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。

**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）

#### <a name="officemailboxenumscomposetype"></a>[MailboxEnums. ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

添加了一个新`ComposeType`枚举，该枚举在撰写模式中可用。

**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）

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

### <a name="sso"></a>SSO

#### <a name="officeruntimeauthgetaccesstoken"></a>[OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
