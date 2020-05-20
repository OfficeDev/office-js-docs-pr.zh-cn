---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序的预览中的功能和 Api。
ms.date: 05/19/2020
localization_priority: Normal
ms.openlocfilehash: 3183c81a9af99f480c2dbecc787695501380cea7
ms.sourcegitcommit: 8499a4247d1cb1e96e99c17cb520f4a8a41667e3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2020
ms.locfileid: "44292292"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="7d04a-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="7d04a-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="7d04a-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="7d04a-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7d04a-105">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="7d04a-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="7d04a-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="7d04a-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="7d04a-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="7d04a-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="7d04a-108">您可以通过[在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)来预览 Web 上 Outlook 中的功能。</span><span class="sxs-lookup"><span data-stu-id="7d04a-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="7d04a-109">此页面上的 "配置预览访问权限" 对适用的功能进行了说明。</span><span class="sxs-lookup"><span data-stu-id="7d04a-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="7d04a-110">对于其他功能，你可以通过填写和提交[此表单](https://aka.ms/OWAPreview)，使用 Microsoft 365 帐户请求对网站上的 Outlook 的预览位的访问权限。</span><span class="sxs-lookup"><span data-stu-id="7d04a-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="7d04a-111">这些功能上记录了 "请求预览访问"。</span><span class="sxs-lookup"><span data-stu-id="7d04a-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="7d04a-112">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="7d04a-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="7d04a-113">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="7d04a-113">Features in preview</span></span>

<span data-ttu-id="7d04a-114">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="7d04a-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="7d04a-115">其他日历属性</span><span class="sxs-lookup"><span data-stu-id="7d04a-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="7d04a-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="7d04a-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="7d04a-117">在撰写模式下添加了一个代表约会全天事件属性的新对象。</span><span class="sxs-lookup"><span data-stu-id="7d04a-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="7d04a-118">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7d04a-118">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="7d04a-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="7d04a-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="7d04a-120">添加了一个新对象，该对象表示在撰写模式下约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="7d04a-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="7d04a-121">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7d04a-121">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="7d04a-122">IsAllDayEvent 的 Office。</span><span class="sxs-lookup"><span data-stu-id="7d04a-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="7d04a-123">添加了一个新属性，该属性表示约会是否为全天事件。</span><span class="sxs-lookup"><span data-stu-id="7d04a-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="7d04a-124">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7d04a-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="7d04a-125">"Context"。项目敏感度</span><span class="sxs-lookup"><span data-stu-id="7d04a-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="7d04a-126">添加了一个表示约会敏感度的新属性。</span><span class="sxs-lookup"><span data-stu-id="7d04a-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="7d04a-127">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7d04a-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="7d04a-128">MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="7d04a-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="7d04a-129">添加了一个 `AppointmentSensitivityType` 代表约会上可用的敏感度选项的新枚举。</span><span class="sxs-lookup"><span data-stu-id="7d04a-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="7d04a-130">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7d04a-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="7d04a-131">发送时追加</span><span class="sxs-lookup"><span data-stu-id="7d04a-131">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="7d04a-132">AppendOnSendAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="7d04a-132">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="7d04a-133">向对象添加了一个新函数 `Body` ，该函数在撰写模式下将数据追加到项正文的末尾。</span><span class="sxs-lookup"><span data-stu-id="7d04a-133">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="7d04a-134">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-134">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="7d04a-135">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="7d04a-135">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="7d04a-136">向清单添加了一个新元素，其中 `AppendOnSend` 扩展权限必须包含在扩展权限的集合中。</span><span class="sxs-lookup"><span data-stu-id="7d04a-136">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="7d04a-137">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="7d04a-138">基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="7d04a-138">Event-based activation</span></span>

<span data-ttu-id="7d04a-139">添加了对 Outlook 外接程序中基于事件的激活功能的支持。若要了解详细信息，请参阅[配置 Outlook 外接程序以进行基于事件的激活](../../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="7d04a-139">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="7d04a-140">LaunchEvent 扩展点</span><span class="sxs-lookup"><span data-stu-id="7d04a-140">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="7d04a-141">`LaunchEvent`向清单添加了扩展点支持。</span><span class="sxs-lookup"><span data-stu-id="7d04a-141">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="7d04a-142">它配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="7d04a-142">It configures event-based activation functionality.</span></span>

<span data-ttu-id="7d04a-143">**适用于**： Outlook 网页版（新式，[请求预览访问](https://aka.ms/OWAPreview)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-143">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="7d04a-144">LaunchEvents 清单元素</span><span class="sxs-lookup"><span data-stu-id="7d04a-144">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="7d04a-145">`LaunchEvents`向清单添加了元素。</span><span class="sxs-lookup"><span data-stu-id="7d04a-145">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="7d04a-146">它支持配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="7d04a-146">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="7d04a-147">**适用于**： Outlook 网页版（新式，[请求预览访问](https://aka.ms/OWAPreview)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-147">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="7d04a-148">运行时清单元素</span><span class="sxs-lookup"><span data-stu-id="7d04a-148">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="7d04a-149">向清单元素添加了 Outlook 支持 `Runtimes` 。</span><span class="sxs-lookup"><span data-stu-id="7d04a-149">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="7d04a-150">它引用了基于事件的激活功能所需的 HTML 和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="7d04a-150">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="7d04a-151">**适用于**： Outlook 网页版（新式，[请求预览访问](https://aka.ms/OWAPreview)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-151">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="7d04a-152">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="7d04a-152">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="7d04a-153">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="7d04a-153">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="7d04a-154">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="7d04a-154">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="7d04a-155">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="7d04a-155">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="7d04a-156">邮件签名</span><span class="sxs-lookup"><span data-stu-id="7d04a-156">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="7d04a-157">SetSignatureAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="7d04a-157">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="7d04a-158">向对象添加了一个新函数 `Body` ，该函数在撰写模式下添加或替换项目正文中的签名。</span><span class="sxs-lookup"><span data-stu-id="7d04a-158">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="7d04a-159">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-159">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="7d04a-160">DisableClientSignatureAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="7d04a-160">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="7d04a-161">添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。</span><span class="sxs-lookup"><span data-stu-id="7d04a-161">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="7d04a-162">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-162">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="7d04a-163">GetComposeTypeAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="7d04a-163">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="7d04a-164">添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。</span><span class="sxs-lookup"><span data-stu-id="7d04a-164">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="7d04a-165">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-165">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="7d04a-166">IsClientSignatureEnabledAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="7d04a-166">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="7d04a-167">添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。</span><span class="sxs-lookup"><span data-stu-id="7d04a-167">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="7d04a-168">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-168">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="7d04a-169">MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="7d04a-169">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="7d04a-170">添加了一个新枚举，该枚举 `ComposeType` 在撰写模式中可用。</span><span class="sxs-lookup"><span data-stu-id="7d04a-170">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="7d04a-171">**适用于**： Windows 上的 outlook （连接到 Office 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="7d04a-171">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="7d04a-172">Office 主题</span><span class="sxs-lookup"><span data-stu-id="7d04a-172">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="7d04a-173">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="7d04a-173">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="7d04a-174">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="7d04a-174">Added ability to get Office theme.</span></span>

<span data-ttu-id="7d04a-175">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7d04a-175">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="7d04a-176">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="7d04a-176">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="7d04a-177">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="7d04a-177">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="7d04a-178">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="7d04a-178">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="7d04a-179">单一登录 (SSO)</span><span class="sxs-lookup"><span data-stu-id="7d04a-179">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="7d04a-180">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="7d04a-180">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="7d04a-181">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="7d04a-181">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="7d04a-182">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="7d04a-182">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="7d04a-183">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7d04a-183">See also</span></span>

- [<span data-ttu-id="7d04a-184">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="7d04a-184">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="7d04a-185">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="7d04a-185">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="7d04a-186">入门</span><span class="sxs-lookup"><span data-stu-id="7d04a-186">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="7d04a-187">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="7d04a-187">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
