---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 10/30/2019
localization_priority: Priority
ms.openlocfilehash: bf8f140e893a19a4dec717b985f3bbf4226db9d5
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902114"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!IMPORTANT]
> 本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

### <a name="integration-with-actionable-messages"></a>与可操作邮件集成

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）

<br>

---

---

### <a name="office-theme"></a>Office 主题

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

增加了获取 Office 主题的功能。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

<br>

---

### <a name="sso"></a>SSO

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

添加了对 `getAccessTokenAsync` 的访问，使外接程序[能够访问](/outlook/add-ins/authenticate-a-user-with-an-sso-token) Microsoft Graph API 的访问令牌。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](/outlook/add-ins/quick-start)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
