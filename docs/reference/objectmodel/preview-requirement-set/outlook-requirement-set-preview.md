---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序和 Office JavaScript Api 的预览中的功能和 Api。
ms.date: 03/27/2020
localization_priority: Normal
ms.openlocfilehash: 3d8eaac1b665d4bd65d5cf0383e53d6f6fb70324
ms.sourcegitcommit: 559a7e178e84947e830cc00dfa01c5c6e398ddc2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2020
ms.locfileid: "43030815"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="068a8-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="068a8-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="068a8-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="068a8-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="068a8-105">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="068a8-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="068a8-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="068a8-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="068a8-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="068a8-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="068a8-108">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="068a8-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="068a8-109">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="068a8-109">Features in preview</span></span>

<span data-ttu-id="068a8-110">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="068a8-110">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="068a8-111">其他日历属性</span><span class="sxs-lookup"><span data-stu-id="068a8-111">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="068a8-112">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="068a8-112">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="068a8-113">在撰写模式下添加了一个代表约会全天事件属性的新对象。</span><span class="sxs-lookup"><span data-stu-id="068a8-113">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="068a8-114">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="068a8-115">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="068a8-115">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="068a8-116">添加了一个新对象，该对象表示在撰写模式下约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="068a8-116">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="068a8-117">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="068a8-118">IsAllDayEvent 的 Office。</span><span class="sxs-lookup"><span data-stu-id="068a8-118">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="068a8-119">添加了一个新属性，该属性表示约会是否为全天事件。</span><span class="sxs-lookup"><span data-stu-id="068a8-119">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="068a8-120">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="068a8-121">"Context"。项目敏感度</span><span class="sxs-lookup"><span data-stu-id="068a8-121">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="068a8-122">添加了一个表示约会敏感度的新属性。</span><span class="sxs-lookup"><span data-stu-id="068a8-122">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="068a8-123">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-123">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="068a8-124">MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="068a8-124">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="068a8-125">添加了一个代表`AppointmentSensitivityType`约会上可用的敏感度选项的新枚举。</span><span class="sxs-lookup"><span data-stu-id="068a8-125">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="068a8-126">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-126">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="068a8-127">发送时追加</span><span class="sxs-lookup"><span data-stu-id="068a8-127">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="068a8-128">AppendOnSendAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="068a8-128">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="068a8-129">向`Body`对象添加了一个新函数，该函数在撰写模式下将数据追加到项正文的末尾。</span><span class="sxs-lookup"><span data-stu-id="068a8-129">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="068a8-130">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="068a8-131">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="068a8-131">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="068a8-132">向清单添加了一个新元素，其中`AppendOnSend`扩展权限必须包含在扩展权限的集合中。</span><span class="sxs-lookup"><span data-stu-id="068a8-132">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="068a8-133">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-133">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="068a8-134">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="068a8-134">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="068a8-135">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="068a8-135">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="068a8-136">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="068a8-136">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="068a8-137">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="068a8-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="068a8-138">邮件签名</span><span class="sxs-lookup"><span data-stu-id="068a8-138">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="068a8-139">SetSignatureAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="068a8-139">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="068a8-140">向`Body`对象添加了一个新函数，该函数在撰写模式下添加或替换项目正文中的签名。</span><span class="sxs-lookup"><span data-stu-id="068a8-140">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="068a8-141">**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="068a8-141">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="068a8-142">DisableClientSignatureAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="068a8-142">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="068a8-143">添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。</span><span class="sxs-lookup"><span data-stu-id="068a8-143">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="068a8-144">**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="068a8-144">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="068a8-145">GetComposeTypeAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="068a8-145">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="068a8-146">添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。</span><span class="sxs-lookup"><span data-stu-id="068a8-146">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="068a8-147">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-147">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="068a8-148">IsClientSignatureEnabledAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="068a8-148">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="068a8-149">添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。</span><span class="sxs-lookup"><span data-stu-id="068a8-149">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="068a8-150">**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="068a8-150">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="068a8-151">MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="068a8-151">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="068a8-152">添加了一个新`ComposeType`枚举，该枚举在撰写模式中可用。</span><span class="sxs-lookup"><span data-stu-id="068a8-152">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="068a8-153">**提供**时间： Windows 上的 outlook （连接到 Office 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="068a8-153">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="068a8-154">Office 主题</span><span class="sxs-lookup"><span data-stu-id="068a8-154">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="068a8-155">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="068a8-155">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="068a8-156">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="068a8-156">Added ability to get Office theme.</span></span>

<span data-ttu-id="068a8-157">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-157">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="068a8-158">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="068a8-158">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="068a8-159">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="068a8-159">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="068a8-160">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="068a8-160">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="sso"></a><span data-ttu-id="068a8-161">SSO</span><span class="sxs-lookup"><span data-stu-id="068a8-161">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="068a8-162">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="068a8-162">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="068a8-163">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="068a8-163">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="068a8-164">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="068a8-164">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="068a8-165">另请参阅</span><span class="sxs-lookup"><span data-stu-id="068a8-165">See also</span></span>

- [<span data-ttu-id="068a8-166">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="068a8-166">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="068a8-167">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="068a8-167">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="068a8-168">入门</span><span class="sxs-lookup"><span data-stu-id="068a8-168">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="068a8-169">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="068a8-169">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
