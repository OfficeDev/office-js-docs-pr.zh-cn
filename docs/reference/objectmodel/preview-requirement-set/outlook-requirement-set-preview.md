---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序的预览中的功能和 Api。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: d91105e0cfbb97dc1a239e40b1c81adc4e76988b
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626594"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="19e27-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="19e27-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="19e27-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="19e27-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="19e27-105">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="19e27-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="19e27-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="19e27-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="19e27-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="19e27-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="19e27-108">您可以通过 [在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)来预览 Web 上 Outlook 中的功能。</span><span class="sxs-lookup"><span data-stu-id="19e27-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="19e27-109">此页面上的 "配置预览访问权限" 对适用的功能进行了说明。</span><span class="sxs-lookup"><span data-stu-id="19e27-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="19e27-110">对于其他功能，你可以通过填写和提交 [此表单](https://aka.ms/OWAPreview)，使用 Microsoft 365 帐户请求对网站上的 Outlook 的预览位的访问权限。</span><span class="sxs-lookup"><span data-stu-id="19e27-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="19e27-111">这些功能上记录了 "请求预览访问"。</span><span class="sxs-lookup"><span data-stu-id="19e27-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="19e27-112">预览要求集包括 [要求集 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md)的所有功能。</span><span class="sxs-lookup"><span data-stu-id="19e27-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="19e27-113">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="19e27-113">Features in preview</span></span>

<span data-ttu-id="19e27-114">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="19e27-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="19e27-115">对受信息权限管理 (IRM) 保护的项的外接程序激活</span><span class="sxs-lookup"><span data-stu-id="19e27-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="19e27-116">现在，外接程序可以在受 IRM 保护的项上激活。</span><span class="sxs-lookup"><span data-stu-id="19e27-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="19e27-117">若要启用此功能，租户管理员需要 `OBJMODEL` 通过在 Office 中设置 " **允许编程访问** " 自定义策略选项来启用使用权限。</span><span class="sxs-lookup"><span data-stu-id="19e27-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="19e27-118">有关详细信息，请参阅 [使用权限和说明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 。</span><span class="sxs-lookup"><span data-stu-id="19e27-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="19e27-119">**适用于**： Windows 中的 Outlook，从内部版本 13229.10000 (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="19e27-120">其他日历属性</span><span class="sxs-lookup"><span data-stu-id="19e27-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="19e27-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="19e27-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="19e27-122">在撰写模式下添加了一个代表约会全天事件属性的新对象。</span><span class="sxs-lookup"><span data-stu-id="19e27-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="19e27-123">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="19e27-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="19e27-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="19e27-125">添加了一个新对象，该对象表示在撰写模式下约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="19e27-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="19e27-126">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="19e27-127">IsAllDayEvent 的 Office。</span><span class="sxs-lookup"><span data-stu-id="19e27-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="19e27-128">添加了一个新属性，该属性表示约会是否为全天事件。</span><span class="sxs-lookup"><span data-stu-id="19e27-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="19e27-129">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="19e27-130">"Context"。项目敏感度</span><span class="sxs-lookup"><span data-stu-id="19e27-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="19e27-131">添加了一个表示约会敏感度的新属性。</span><span class="sxs-lookup"><span data-stu-id="19e27-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="19e27-132">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="19e27-133">MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="19e27-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="19e27-134">添加了一个 `AppointmentSensitivityType` 代表约会上可用的敏感度选项的新枚举。</span><span class="sxs-lookup"><span data-stu-id="19e27-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="19e27-135">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="19e27-136">基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="19e27-136">Event-based activation</span></span>

<span data-ttu-id="19e27-137">添加了对 Outlook 外接程序中基于事件的激活功能的支持。若要了解详细信息，请参阅 [配置 Outlook 外接程序以进行基于事件的激活](../../../outlook/autolaunch.md) 。</span><span class="sxs-lookup"><span data-stu-id="19e27-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="19e27-138">LaunchEvent 扩展点</span><span class="sxs-lookup"><span data-stu-id="19e27-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="19e27-139">`LaunchEvent`向清单添加了扩展点支持。</span><span class="sxs-lookup"><span data-stu-id="19e27-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="19e27-140">它配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="19e27-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="19e27-141">**中的可用**： Outlook 网页版 (新式， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="19e27-141">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="19e27-142">LaunchEvents 清单元素</span><span class="sxs-lookup"><span data-stu-id="19e27-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="19e27-143">`LaunchEvents`向清单添加了元素。</span><span class="sxs-lookup"><span data-stu-id="19e27-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="19e27-144">它支持配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="19e27-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="19e27-145">**中的可用**： Outlook 网页版 (新式， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="19e27-145">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="19e27-146">运行时清单元素</span><span class="sxs-lookup"><span data-stu-id="19e27-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="19e27-147">向清单元素添加了 Outlook 支持 `Runtimes` 。</span><span class="sxs-lookup"><span data-stu-id="19e27-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="19e27-148">它引用了基于事件的激活功能所需的 HTML 和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="19e27-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="19e27-149">**中的可用**： Outlook 网页版 (新式， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="19e27-149">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="19e27-150">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="19e27-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="19e27-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="19e27-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="19e27-152">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="19e27-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="19e27-153">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (经典) </span><span class="sxs-lookup"><span data-stu-id="19e27-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="19e27-154">邮件签名</span><span class="sxs-lookup"><span data-stu-id="19e27-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="19e27-155">SetSignatureAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="19e27-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="19e27-156">向对象添加了一个新函数 `Body` ，该函数在撰写模式下添加或替换项目正文中的签名。</span><span class="sxs-lookup"><span data-stu-id="19e27-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="19e27-157">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="19e27-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="19e27-158">DisableClientSignatureAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="19e27-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="19e27-159">添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。</span><span class="sxs-lookup"><span data-stu-id="19e27-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="19e27-160">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="19e27-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="19e27-161">GetComposeTypeAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="19e27-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="19e27-162">添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。</span><span class="sxs-lookup"><span data-stu-id="19e27-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="19e27-163">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="19e27-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="19e27-164">IsClientSignatureEnabledAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="19e27-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="19e27-165">添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。</span><span class="sxs-lookup"><span data-stu-id="19e27-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="19e27-166">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="19e27-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="19e27-167">MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="19e27-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="19e27-168">添加了一个新枚举，该枚举 `ComposeType` 在撰写模式中可用。</span><span class="sxs-lookup"><span data-stu-id="19e27-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="19e27-169">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="19e27-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="19e27-170">包含操作的通知邮件</span><span class="sxs-lookup"><span data-stu-id="19e27-170">Notification messages with actions</span></span>

<span data-ttu-id="19e27-171">通过此功能，您的外接程序可以在默认 **取消** 操作之外包含具有自定义操作的通知消息。</span><span class="sxs-lookup"><span data-stu-id="19e27-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="19e27-172">NotificationMessageDetails</span><span class="sxs-lookup"><span data-stu-id="19e27-172">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="19e27-173">添加了一个新属性，您可以 `InsightMessage` 使用自定义操作添加通知。</span><span class="sxs-lookup"><span data-stu-id="19e27-173">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="19e27-174">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="19e27-174">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="19e27-175">NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="19e27-175">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="19e27-176">添加了一个新对象，可在其中为通知定义自定义操作 `InsightMessage` 。</span><span class="sxs-lookup"><span data-stu-id="19e27-176">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="19e27-177">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="19e27-177">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="19e27-178">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="19e27-178">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="19e27-179">添加了新枚举 `ActionType` 。</span><span class="sxs-lookup"><span data-stu-id="19e27-179">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="19e27-180">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="19e27-180">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="19e27-181">ItemNotificationMessageType InsightMessage</span><span class="sxs-lookup"><span data-stu-id="19e27-181">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="19e27-182">向枚举添加了新类型 `InsightMessage` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="19e27-182">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="19e27-183">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="19e27-183">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="19e27-184">Office 主题</span><span class="sxs-lookup"><span data-stu-id="19e27-184">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="19e27-185">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="19e27-185">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="19e27-186">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="19e27-186">Added ability to get Office theme.</span></span>

<span data-ttu-id="19e27-187">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-187">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="19e27-188">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="19e27-188">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="19e27-189">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="19e27-189">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="19e27-190">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-190">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="19e27-191">会话数据</span><span class="sxs-lookup"><span data-stu-id="19e27-191">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="19e27-192">SessionData</span><span class="sxs-lookup"><span data-stu-id="19e27-192">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="19e27-193">添加了一个代表项目的会话数据的新对象。</span><span class="sxs-lookup"><span data-stu-id="19e27-193">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="19e27-194">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-194">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="19e27-195">SessionData 的 Office。</span><span class="sxs-lookup"><span data-stu-id="19e27-195">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="19e27-196">添加了一个新属性以在撰写模式下管理项目的会话数据。</span><span class="sxs-lookup"><span data-stu-id="19e27-196">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="19e27-197">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="19e27-197">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

## <a name="see-also"></a><span data-ttu-id="19e27-198">另请参阅</span><span class="sxs-lookup"><span data-stu-id="19e27-198">See also</span></span>

- [<span data-ttu-id="19e27-199">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="19e27-199">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="19e27-200">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="19e27-200">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="19e27-201">入门</span><span class="sxs-lookup"><span data-stu-id="19e27-201">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="19e27-202">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="19e27-202">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
