---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序的预览中的功能和 Api。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: c2b4d31fdb545afdc695c5aef84856aeaebdbf28
ms.sourcegitcommit: b634bfe9a946fbd95754e87f070a904ed57586ff
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/15/2020
ms.locfileid: "44253626"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="5087f-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="5087f-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="5087f-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="5087f-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5087f-105">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="5087f-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="5087f-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="5087f-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="5087f-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="5087f-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="5087f-108">您可以通过[在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)来预览 Web 上 Outlook 中的功能。</span><span class="sxs-lookup"><span data-stu-id="5087f-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="5087f-109">此页面上的 "配置预览访问权限" 对适用的功能进行了说明。</span><span class="sxs-lookup"><span data-stu-id="5087f-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="5087f-110">对于其他功能，你可以通过填写和提交[此表单](https://aka.ms/OWAPreview)，使用 Microsoft 365 帐户请求对网站上的 Outlook 的预览位的访问权限。</span><span class="sxs-lookup"><span data-stu-id="5087f-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="5087f-111">这些功能上注明了 "请求访问"。</span><span class="sxs-lookup"><span data-stu-id="5087f-111">"Request access" is noted on those features.</span></span>

<span data-ttu-id="5087f-112">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="5087f-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="5087f-113">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="5087f-113">Features in preview</span></span>

<span data-ttu-id="5087f-114">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="5087f-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="5087f-115">其他日历属性</span><span class="sxs-lookup"><span data-stu-id="5087f-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="5087f-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="5087f-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="5087f-117">在撰写模式下添加了一个代表约会全天事件属性的新对象。</span><span class="sxs-lookup"><span data-stu-id="5087f-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="5087f-118">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5087f-118">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="5087f-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="5087f-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="5087f-120">添加了一个新对象，该对象表示在撰写模式下约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="5087f-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="5087f-121">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5087f-121">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="5087f-122">IsAllDayEvent 的 Office。</span><span class="sxs-lookup"><span data-stu-id="5087f-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="5087f-123">添加了一个新属性，该属性表示约会是否为全天事件。</span><span class="sxs-lookup"><span data-stu-id="5087f-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="5087f-124">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5087f-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="5087f-125">"Context"。项目敏感度</span><span class="sxs-lookup"><span data-stu-id="5087f-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="5087f-126">添加了一个表示约会敏感度的新属性。</span><span class="sxs-lookup"><span data-stu-id="5087f-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="5087f-127">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5087f-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="5087f-128">MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="5087f-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="5087f-129">添加了一个 `AppointmentSensitivityType` 代表约会上可用的敏感度选项的新枚举。</span><span class="sxs-lookup"><span data-stu-id="5087f-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="5087f-130">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5087f-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="5087f-131">发送时追加</span><span class="sxs-lookup"><span data-stu-id="5087f-131">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="5087f-132">AppendOnSendAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="5087f-132">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="5087f-133">向对象添加了一个新函数 `Body` ，该函数在撰写模式下将数据追加到项正文的末尾。</span><span class="sxs-lookup"><span data-stu-id="5087f-133">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="5087f-134">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="5087f-134">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="5087f-135">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="5087f-135">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="5087f-136">向清单添加了一个新元素，其中 `AppendOnSend` 扩展权限必须包含在扩展权限的集合中。</span><span class="sxs-lookup"><span data-stu-id="5087f-136">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="5087f-137">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="5087f-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="5087f-138">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="5087f-138">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="5087f-139">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="5087f-139">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="5087f-140">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="5087f-140">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="5087f-141">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="5087f-141">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="5087f-142">邮件签名</span><span class="sxs-lookup"><span data-stu-id="5087f-142">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="5087f-143">SetSignatureAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="5087f-143">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="5087f-144">向对象添加了一个新函数 `Body` ，该函数在撰写模式下添加或替换项目正文中的签名。</span><span class="sxs-lookup"><span data-stu-id="5087f-144">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="5087f-145">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="5087f-145">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="5087f-146">DisableClientSignatureAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="5087f-146">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="5087f-147">添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。</span><span class="sxs-lookup"><span data-stu-id="5087f-147">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="5087f-148">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="5087f-148">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="5087f-149">GetComposeTypeAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="5087f-149">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="5087f-150">添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。</span><span class="sxs-lookup"><span data-stu-id="5087f-150">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="5087f-151">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="5087f-151">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="5087f-152">IsClientSignatureEnabledAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="5087f-152">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="5087f-153">添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。</span><span class="sxs-lookup"><span data-stu-id="5087f-153">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="5087f-154">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="5087f-154">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="5087f-155">MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="5087f-155">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="5087f-156">添加了一个新枚举，该枚举 `ComposeType` 在撰写模式中可用。</span><span class="sxs-lookup"><span data-stu-id="5087f-156">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="5087f-157">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="5087f-157">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="5087f-158">Office 主题</span><span class="sxs-lookup"><span data-stu-id="5087f-158">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="5087f-159">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="5087f-159">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="5087f-160">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="5087f-160">Added ability to get Office theme.</span></span>

<span data-ttu-id="5087f-161">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5087f-161">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="5087f-162">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="5087f-162">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="5087f-163">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="5087f-163">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="5087f-164">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5087f-164">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="online-meeting-provider-integration"></a><span data-ttu-id="5087f-165">联机会议提供程序集成</span><span class="sxs-lookup"><span data-stu-id="5087f-165">Online meeting provider integration</span></span>

<span data-ttu-id="5087f-166">添加了对 Outlook 移动外接程序中的联机会议集成的支持。有关详细信息，请参阅为[联机会议提供商创建 Outlook mobile 外接程序](../../../outlook/online-meeting.md)以了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="5087f-166">Added support for online-meeting integration in Outlook mobile add-ins. See [Create an Outlook mobile add-in for an online-meeting provider](../../../outlook/online-meeting.md) to learn more.</span></span>

#### <a name="mobileonlinemeetingcommandsurface-extension-point"></a>[<span data-ttu-id="5087f-167">MobileOnlineMeetingCommandSurface 扩展点</span><span class="sxs-lookup"><span data-stu-id="5087f-167">MobileOnlineMeetingCommandSurface extension point</span></span>](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

<span data-ttu-id="5087f-168">`MobileOnlineMeetingCommandSurface`向清单添加了扩展点。</span><span class="sxs-lookup"><span data-stu-id="5087f-168">Added `MobileOnlineMeetingCommandSurface` extension point to manifest.</span></span> <span data-ttu-id="5087f-169">它定义联机会议集成。</span><span class="sxs-lookup"><span data-stu-id="5087f-169">It defines the online meeting integration.</span></span>

<span data-ttu-id="5087f-170">**适用于**： Outlook on Android （连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5087f-170">**Available in**: Outlook on Android (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="sso"></a><span data-ttu-id="5087f-171">SSO</span><span class="sxs-lookup"><span data-stu-id="5087f-171">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="5087f-172">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="5087f-172">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="5087f-173">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="5087f-173">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="5087f-174">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="5087f-174">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="5087f-175">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5087f-175">See also</span></span>

- [<span data-ttu-id="5087f-176">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="5087f-176">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="5087f-177">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="5087f-177">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="5087f-178">入门</span><span class="sxs-lookup"><span data-stu-id="5087f-178">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="5087f-179">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="5087f-179">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
