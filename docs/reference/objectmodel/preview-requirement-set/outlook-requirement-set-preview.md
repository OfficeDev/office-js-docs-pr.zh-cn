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
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="1ad92-102">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="1ad92-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="1ad92-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="1ad92-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1ad92-104">本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="1ad92-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="1ad92-105">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="1ad92-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="1ad92-106">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="1ad92-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="1ad92-107">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="1ad92-107">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="1ad92-108">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="1ad92-108">Features in preview</span></span>

<span data-ttu-id="1ad92-109">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="1ad92-109">The following features are in preview.</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="1ad92-110">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="1ad92-110">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="1ad92-111">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="1ad92-111">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="1ad92-112">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="1ad92-112">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="1ad92-113">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="1ad92-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="1ad92-114">Office 主题</span><span class="sxs-lookup"><span data-stu-id="1ad92-114">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="1ad92-115">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="1ad92-115">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="1ad92-116">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="1ad92-116">Added ability to get Office theme.</span></span>

<span data-ttu-id="1ad92-117">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="1ad92-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="1ad92-118">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="1ad92-118">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="1ad92-119">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="1ad92-119">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="1ad92-120">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="1ad92-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="1ad92-121">SSO</span><span class="sxs-lookup"><span data-stu-id="1ad92-121">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="1ad92-122">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1ad92-122">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="1ad92-123">添加了对 `getAccessTokenAsync` 的访问，使外接程序[能够访问](/outlook/add-ins/authenticate-a-user-with-an-sso-token) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="1ad92-123">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="1ad92-124">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="1ad92-124">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="1ad92-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1ad92-125">See also</span></span>

- [<span data-ttu-id="1ad92-126">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="1ad92-126">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="1ad92-127">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="1ad92-127">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="1ad92-128">入门</span><span class="sxs-lookup"><span data-stu-id="1ad92-128">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="1ad92-129">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="1ad92-129">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
