---
title: Outlook 外接程序 API 预览要求集
description: Outlook 外接程序当前处于预览阶段的功能和 API。
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: 39dd1221f4dea9674c89cdaad20024ce408f8db3
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104838"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="74e47-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="74e47-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="74e47-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="74e47-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="74e47-105">本文档适用于 **预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="74e47-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="74e47-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="74e47-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="74e47-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="74e47-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="74e47-108">你可能能够通过在 [Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)租户上配置定向版本来预览 Outlook 网页版中的功能。</span><span class="sxs-lookup"><span data-stu-id="74e47-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="74e47-109">"配置预览访问"在此页面上针对适用的功能进行说明。</span><span class="sxs-lookup"><span data-stu-id="74e47-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="74e47-110">对于其他功能，你可能能够通过完成和提交此表单，请求访问 Outlook 网页版预览位，使用 Microsoft 365 [帐户](https://aka.ms/OWAPreview)。</span><span class="sxs-lookup"><span data-stu-id="74e47-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="74e47-111">这些功能上会指出"请求预览访问"。</span><span class="sxs-lookup"><span data-stu-id="74e47-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="74e47-112">预览要求集包括要求集 [1.9 的所有功能](../requirement-set-1.9/outlook-requirement-set-1.9.md)。</span><span class="sxs-lookup"><span data-stu-id="74e47-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="74e47-113">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="74e47-113">Features in preview</span></span>

<span data-ttu-id="74e47-114">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="74e47-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="74e47-115">对受信息权限管理中心 IRM 保护的项目 (加载项) </span><span class="sxs-lookup"><span data-stu-id="74e47-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="74e47-116">加载项现在可以在受 IRM 保护的项目上激活。</span><span class="sxs-lookup"><span data-stu-id="74e47-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="74e47-117">若要启用此功能，租户管理员需要通过设置 Office 中的"允许编程访问自定义策略"选项来 `OBJMODEL` 启用使用权限。 </span><span class="sxs-lookup"><span data-stu-id="74e47-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="74e47-118">有关详细信息 [，请参阅使用](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 权限和说明。</span><span class="sxs-lookup"><span data-stu-id="74e47-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="74e47-119">**适用于：Windows** 版 Outlook，从内部版本 13229.10000 (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="74e47-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="74e47-120">其他日历属性</span><span class="sxs-lookup"><span data-stu-id="74e47-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="74e47-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="74e47-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="74e47-122">新增了一个对象，该对象代表撰写模式下约会的全天事件属性。</span><span class="sxs-lookup"><span data-stu-id="74e47-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="74e47-123">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="74e47-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="74e47-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="74e47-125">新增了一个对象，该对象代表撰写模式下约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="74e47-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="74e47-126">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="74e47-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="74e47-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="74e47-128">添加了一个新属性，表示约会是全天事件。</span><span class="sxs-lookup"><span data-stu-id="74e47-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="74e47-129">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="74e47-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="74e47-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="74e47-131">添加了一个新属性，表示约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="74e47-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="74e47-132">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="74e47-133">Office.MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="74e47-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="74e47-134">添加了一个新枚举 `AppointmentSensitivityType` ，该枚举代表约会可用的敏感度选项。</span><span class="sxs-lookup"><span data-stu-id="74e47-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="74e47-135">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="74e47-136">基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="74e47-136">Event-based activation</span></span>

<span data-ttu-id="74e47-137">增加了对 Outlook 外接程序中基于事件的激活功能的支持。请参阅 [配置 Outlook 外接程序进行基于事件的激活](../../../outlook/autolaunch.md) 以了解更多信息。</span><span class="sxs-lookup"><span data-stu-id="74e47-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="74e47-138">LaunchEvent 扩展点</span><span class="sxs-lookup"><span data-stu-id="74e47-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="74e47-139">向 `LaunchEvent` 清单添加了扩展点支持。</span><span class="sxs-lookup"><span data-stu-id="74e47-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="74e47-140">它配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="74e47-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="74e47-141">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="74e47-141">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="74e47-142">LaunchEvents 清单元素</span><span class="sxs-lookup"><span data-stu-id="74e47-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="74e47-143">向 `LaunchEvents` 清单添加了元素。</span><span class="sxs-lookup"><span data-stu-id="74e47-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="74e47-144">它支持配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="74e47-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="74e47-145">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="74e47-145">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="74e47-146">运行时清单元素</span><span class="sxs-lookup"><span data-stu-id="74e47-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="74e47-147">向清单元素添加了 Outlook `Runtimes` 支持。</span><span class="sxs-lookup"><span data-stu-id="74e47-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="74e47-148">它引用基于事件的激活功能所需的 HTML 和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="74e47-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="74e47-149">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="74e47-149">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="74e47-150">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="74e47-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="74e47-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="74e47-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="74e47-152">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="74e47-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="74e47-153">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="74e47-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="74e47-154">邮件签名</span><span class="sxs-lookup"><span data-stu-id="74e47-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="74e47-155">Office.context.mailbox.item.body.setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="74e47-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="74e47-156">向对象添加了一个新函数，该函数在撰写模式下添加或替换 `Body` 项目正文中的签名。</span><span class="sxs-lookup"><span data-stu-id="74e47-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="74e47-157">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="74e47-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="74e47-158">Office.context.mailbox.item.disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="74e47-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="74e47-159">添加了一个新函数，该函数在撰写模式下禁用发送邮箱的客户端签名。</span><span class="sxs-lookup"><span data-stu-id="74e47-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="74e47-160">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="74e47-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="74e47-161">Office.context.mailbox.item.getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="74e47-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="74e47-162">添加了一个新函数，该函数获取撰写模式下邮件的撰写类型。</span><span class="sxs-lookup"><span data-stu-id="74e47-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="74e47-163">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="74e47-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="74e47-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="74e47-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="74e47-165">新增了一个函数，该函数检查在撰写模式下是否对项目启用了客户端签名。</span><span class="sxs-lookup"><span data-stu-id="74e47-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="74e47-166">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="74e47-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="74e47-167">Office.MailboxEnums.ComposeType</span><span class="sxs-lookup"><span data-stu-id="74e47-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="74e47-168">添加了在撰写模式下 `ComposeType` 可用的新枚举。</span><span class="sxs-lookup"><span data-stu-id="74e47-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="74e47-169">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式、配置预览) [](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="74e47-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="74e47-170">包含操作的通知邮件</span><span class="sxs-lookup"><span data-stu-id="74e47-170">Notification messages with actions</span></span>

<span data-ttu-id="74e47-171">此功能允许外接程序在默认消除操作之外包含包含自定义操作 **的通知** 消息。</span><span class="sxs-lookup"><span data-stu-id="74e47-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span> <span data-ttu-id="74e47-172">在新式 Outlook 网页中，此功能仅在撰写模式下可用。</span><span class="sxs-lookup"><span data-stu-id="74e47-172">In modern Outlook on the web, this feature is available in Compose mode only.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="74e47-173">Office.NotificationMessageDetails.actions</span><span class="sxs-lookup"><span data-stu-id="74e47-173">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="74e47-174">添加了一个新属性，允许您使用自定义操作 `InsightMessage` 添加通知。</span><span class="sxs-lookup"><span data-stu-id="74e47-174">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="74e47-175">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="74e47-175">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="74e47-176">Office.NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="74e47-176">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="74e47-177">添加了一个新对象，用于定义通知的自定义 `InsightMessage` 操作。</span><span class="sxs-lookup"><span data-stu-id="74e47-177">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="74e47-178">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="74e47-178">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="74e47-179">Office.MailboxEnums.ActionType</span><span class="sxs-lookup"><span data-stu-id="74e47-179">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="74e47-180">添加了一个新枚举 `ActionType` 。</span><span class="sxs-lookup"><span data-stu-id="74e47-180">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="74e47-181">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="74e47-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="74e47-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span><span class="sxs-lookup"><span data-stu-id="74e47-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="74e47-183">向枚举添加了 `InsightMessage` 一个新 `ItemNotificationMessageType` 类型。</span><span class="sxs-lookup"><span data-stu-id="74e47-183">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="74e47-184">**适用于：Windows** 版 Outlook (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="74e47-184">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="74e47-185">Office 主题</span><span class="sxs-lookup"><span data-stu-id="74e47-185">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="74e47-186">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="74e47-186">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="74e47-187">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="74e47-187">Added ability to get Office theme.</span></span>

<span data-ttu-id="74e47-188">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-188">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="74e47-189">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="74e47-189">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="74e47-190">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="74e47-190">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="74e47-191">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-191">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="74e47-192">会话数据</span><span class="sxs-lookup"><span data-stu-id="74e47-192">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="74e47-193">Office.SessionData</span><span class="sxs-lookup"><span data-stu-id="74e47-193">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="74e47-194">添加了一个新对象，该对象代表项目的会话数据。</span><span class="sxs-lookup"><span data-stu-id="74e47-194">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="74e47-195">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="74e47-196">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="74e47-196">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="74e47-197">添加了一个新属性，用于管理撰写模式下项目的会话数据。</span><span class="sxs-lookup"><span data-stu-id="74e47-197">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="74e47-198">**适用于：** 连接到 Microsoft 365 (Windows 版 Outlook) </span><span class="sxs-lookup"><span data-stu-id="74e47-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

## <a name="see-also"></a><span data-ttu-id="74e47-199">另请参阅</span><span class="sxs-lookup"><span data-stu-id="74e47-199">See also</span></span>

- [<span data-ttu-id="74e47-200">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="74e47-200">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="74e47-201">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="74e47-201">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="74e47-202">入门</span><span class="sxs-lookup"><span data-stu-id="74e47-202">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="74e47-203">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="74e47-203">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
