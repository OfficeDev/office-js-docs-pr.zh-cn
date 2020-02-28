---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 87c15ac889a955412e6a8350baaed8611fdb5164
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325218"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="7faab-102">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="7faab-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="7faab-103">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="7faab-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7faab-104">本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="7faab-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="7faab-105">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="7faab-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="7faab-106">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="7faab-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="7faab-107">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="7faab-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="7faab-108">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="7faab-108">Features in preview</span></span>

<span data-ttu-id="7faab-109">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="7faab-109">The following features are in preview.</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="7faab-110">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="7faab-110">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="7faab-111">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="7faab-111">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="7faab-112">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="7faab-112">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="7faab-113">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="7faab-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="7faab-114">Office 主题</span><span class="sxs-lookup"><span data-stu-id="7faab-114">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="7faab-115">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="7faab-115">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="7faab-116">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="7faab-116">Added ability to get Office theme.</span></span>

<span data-ttu-id="7faab-117">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7faab-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="7faab-118">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="7faab-118">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="7faab-119">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="7faab-119">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="7faab-120">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7faab-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="7faab-121">SSO</span><span class="sxs-lookup"><span data-stu-id="7faab-121">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="7faab-122">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="7faab-122">OfficeRuntime.auth.getAccessToken</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="7faab-123">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="7faab-123">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="7faab-124">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="7faab-124">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="7faab-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7faab-125">See also</span></span>

- [<span data-ttu-id="7faab-126">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="7faab-126">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="7faab-127">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="7faab-127">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="7faab-128">入门</span><span class="sxs-lookup"><span data-stu-id="7faab-128">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="7faab-129">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="7faab-129">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
