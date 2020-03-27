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
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="6f0b6-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="6f0b6-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="6f0b6-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6f0b6-105">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="6f0b6-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="6f0b6-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="6f0b6-108">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="6f0b6-109">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="6f0b6-109">Features in preview</span></span>

<span data-ttu-id="6f0b6-110">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-110">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="6f0b6-111">发送时追加</span><span class="sxs-lookup"><span data-stu-id="6f0b6-111">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="6f0b6-112">AppendOnSendAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="6f0b6-112">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="6f0b6-113">向`Body`对象添加了一个新函数，该函数在撰写模式下将数据追加到项正文的末尾。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-113">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="6f0b6-114">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="6f0b6-115">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="6f0b6-115">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="6f0b6-116">向清单添加了一个新元素，其中`AppendOnSend`扩展权限必须包含在扩展权限的集合中。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-116">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="6f0b6-117">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="6f0b6-118">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="6f0b6-118">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="6f0b6-119">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="6f0b6-119">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6f0b6-120">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-120">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="6f0b6-121">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-121">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="6f0b6-122">邮件签名</span><span class="sxs-lookup"><span data-stu-id="6f0b6-122">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="6f0b6-123">SetSignatureAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="6f0b6-123">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="6f0b6-124">向`Body`对象添加了一个新函数，该函数在撰写模式下添加或替换项目正文中的签名。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-124">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="6f0b6-125">**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-125">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="6f0b6-126">DisableClientSignatureAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-126">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6f0b6-127">添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-127">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="6f0b6-128">**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-128">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="6f0b6-129">GetComposeTypeAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-129">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="6f0b6-130">添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-130">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="6f0b6-131">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-131">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="6f0b6-132">IsClientSignatureEnabledAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-132">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6f0b6-133">添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-133">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="6f0b6-134">**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-134">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="6f0b6-135">MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="6f0b6-135">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="6f0b6-136">添加了一个新`ComposeType`枚举，该枚举在撰写模式中可用。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-136">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="6f0b6-137">**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="6f0b6-138">Office 主题</span><span class="sxs-lookup"><span data-stu-id="6f0b6-138">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="6f0b6-139">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="6f0b6-139">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="6f0b6-140">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-140">Added ability to get Office theme.</span></span>

<span data-ttu-id="6f0b6-141">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-141">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="6f0b6-142">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="6f0b6-142">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="6f0b6-143">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-143">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="6f0b6-144">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-144">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="sso"></a><span data-ttu-id="6f0b6-145">SSO</span><span class="sxs-lookup"><span data-stu-id="6f0b6-145">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="6f0b6-146">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="6f0b6-146">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="6f0b6-147">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="6f0b6-147">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="6f0b6-148">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="6f0b6-148">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="6f0b6-149">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6f0b6-149">See also</span></span>

- [<span data-ttu-id="6f0b6-150">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="6f0b6-150">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="6f0b6-151">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="6f0b6-151">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="6f0b6-152">入门</span><span class="sxs-lookup"><span data-stu-id="6f0b6-152">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="6f0b6-153">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="6f0b6-153">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
