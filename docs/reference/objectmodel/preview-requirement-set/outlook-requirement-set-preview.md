---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序和 Office JavaScript Api 的预览中的功能和 Api。
ms.date: 03/04/2020
localization_priority: Normal
ms.openlocfilehash: c87ce8472becc072702f58e7d8c21665904673d2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717808"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="9255e-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="9255e-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="9255e-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="9255e-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9255e-105">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="9255e-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="9255e-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="9255e-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="9255e-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="9255e-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="9255e-108">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="9255e-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="9255e-109">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="9255e-109">Features in preview</span></span>

<span data-ttu-id="9255e-110">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="9255e-110">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="9255e-111">发送时追加</span><span class="sxs-lookup"><span data-stu-id="9255e-111">Append on send</span></span>

#### <a name="officebodyappendonsendasync"></a>[<span data-ttu-id="9255e-112">AppendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="9255e-112">Office.Body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="9255e-113">向`Body`对象添加了一个新函数，该函数在撰写模式下将数据追加到项正文的末尾。</span><span class="sxs-lookup"><span data-stu-id="9255e-113">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="9255e-114">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="9255e-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="9255e-115">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="9255e-115">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="9255e-116">向清单添加了一个新元素，其中`AppendOnSend`扩展权限必须包含在扩展权限的集合中。</span><span class="sxs-lookup"><span data-stu-id="9255e-116">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="9255e-117">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="9255e-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="9255e-118">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="9255e-118">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="9255e-119">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="9255e-119">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9255e-120">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="9255e-120">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="9255e-121">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="9255e-121">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="9255e-122">Office 主题</span><span class="sxs-lookup"><span data-stu-id="9255e-122">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="9255e-123">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="9255e-123">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="9255e-124">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="9255e-124">Added ability to get Office theme.</span></span>

<span data-ttu-id="9255e-125">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="9255e-125">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="9255e-126">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="9255e-126">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9255e-127">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="9255e-127">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="9255e-128">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="9255e-128">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="9255e-129">SSO</span><span class="sxs-lookup"><span data-stu-id="9255e-129">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="9255e-130">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="9255e-130">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="9255e-131">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="9255e-131">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="9255e-132">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="9255e-132">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="9255e-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9255e-133">See also</span></span>

- [<span data-ttu-id="9255e-134">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="9255e-134">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="9255e-135">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="9255e-135">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9255e-136">入门</span><span class="sxs-lookup"><span data-stu-id="9255e-136">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="9255e-137">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="9255e-137">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
