---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序的预览中的功能和 Api。
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 24cad394f0f3ffb95a05a81ccb38ee4aa72a3797
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431064"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="eadeb-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="eadeb-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="eadeb-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="eadeb-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="eadeb-105">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="eadeb-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="eadeb-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="eadeb-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="eadeb-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="eadeb-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="eadeb-108">您可以通过 [在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)来预览 Web 上 Outlook 中的功能。</span><span class="sxs-lookup"><span data-stu-id="eadeb-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="eadeb-109">此页面上的 "配置预览访问权限" 对适用的功能进行了说明。</span><span class="sxs-lookup"><span data-stu-id="eadeb-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="eadeb-110">对于其他功能，你可以通过填写和提交 [此表单](https://aka.ms/OWAPreview)，使用 Microsoft 365 帐户请求对网站上的 Outlook 的预览位的访问权限。</span><span class="sxs-lookup"><span data-stu-id="eadeb-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="eadeb-111">这些功能上记录了 "请求预览访问"。</span><span class="sxs-lookup"><span data-stu-id="eadeb-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="eadeb-112">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="eadeb-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="eadeb-113">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="eadeb-113">Features in preview</span></span>

<span data-ttu-id="eadeb-114">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="eadeb-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="eadeb-115">对受信息权限管理 (IRM) 保护的项的外接程序激活</span><span class="sxs-lookup"><span data-stu-id="eadeb-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="eadeb-116">现在，外接程序可以在受 IRM 保护的项上激活。</span><span class="sxs-lookup"><span data-stu-id="eadeb-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="eadeb-117">若要启用此功能，租户管理员需要 `OBJMODEL` 通过在 Office 中设置 " **允许编程访问** " 自定义策略选项来启用使用权限。</span><span class="sxs-lookup"><span data-stu-id="eadeb-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="eadeb-118">有关详细信息，请参阅 [使用权限和说明](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="eadeb-119">**适用于**： Windows 中的 Outlook，从内部版本 13229.10000 (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="eadeb-120">其他日历属性</span><span class="sxs-lookup"><span data-stu-id="eadeb-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="eadeb-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="eadeb-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="eadeb-122">在撰写模式下添加了一个代表约会全天事件属性的新对象。</span><span class="sxs-lookup"><span data-stu-id="eadeb-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="eadeb-123">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="eadeb-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="eadeb-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="eadeb-125">添加了一个新对象，该对象表示在撰写模式下约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="eadeb-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="eadeb-126">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="eadeb-127">IsAllDayEvent 的 Office。</span><span class="sxs-lookup"><span data-stu-id="eadeb-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="eadeb-128">添加了一个新属性，该属性表示约会是否为全天事件。</span><span class="sxs-lookup"><span data-stu-id="eadeb-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="eadeb-129">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="eadeb-130">"Context"。项目敏感度</span><span class="sxs-lookup"><span data-stu-id="eadeb-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="eadeb-131">添加了一个表示约会敏感度的新属性。</span><span class="sxs-lookup"><span data-stu-id="eadeb-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="eadeb-132">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="eadeb-133">MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="eadeb-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="eadeb-134">添加了一个 `AppointmentSensitivityType` 代表约会上可用的敏感度选项的新枚举。</span><span class="sxs-lookup"><span data-stu-id="eadeb-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="eadeb-135">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="eadeb-136">发送时附加</span><span class="sxs-lookup"><span data-stu-id="eadeb-136">Append on send</span></span>

<span data-ttu-id="eadeb-137">若要了解如何使用 "发送时追加" 功能，请参阅在 [Outlook 加载项中实施 "在发送时实现附加](../../../outlook/append-on-send.md)"。</span><span class="sxs-lookup"><span data-stu-id="eadeb-137">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="eadeb-138">AppendOnSendAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="eadeb-138">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-)

<span data-ttu-id="eadeb-139">向对象添加了一个新函数 `Body` ，该函数在撰写模式下将数据追加到项正文的末尾。</span><span class="sxs-lookup"><span data-stu-id="eadeb-139">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="eadeb-140">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="eadeb-141">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="eadeb-141">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="eadeb-142">向清单添加了一个新元素，其中 `AppendOnSend` 扩展权限必须包含在扩展权限的集合中。</span><span class="sxs-lookup"><span data-stu-id="eadeb-142">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="eadeb-143">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="eadeb-144">Api 的异步版本 `display`</span><span class="sxs-lookup"><span data-stu-id="eadeb-144">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="eadeb-145">DisplayAppointmentFormAsync 的</span><span class="sxs-lookup"><span data-stu-id="eadeb-145">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="eadeb-146">向显示现有约会的对象添加了新函数 `Mailbox` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-146">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="eadeb-147">这是方法的异步版本 `displayAppointmentForm` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-147">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="eadeb-148">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="eadeb-149">DisplayMessageFormAsync 的</span><span class="sxs-lookup"><span data-stu-id="eadeb-149">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="eadeb-150">向显示现有邮件的对象添加了新函数 `Mailbox` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-150">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="eadeb-151">这是方法的异步版本 `displayMessageForm` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-151">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="eadeb-152">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-152">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="eadeb-153">DisplayNewAppointmentFormAsync 的</span><span class="sxs-lookup"><span data-stu-id="eadeb-153">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="eadeb-154">向 `Mailbox` 显示新约会窗体的对象添加了一个新函数。</span><span class="sxs-lookup"><span data-stu-id="eadeb-154">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="eadeb-155">这是方法的异步版本 `displayNewAppointmentForm` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-155">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="eadeb-156">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-156">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="eadeb-157">DisplayNewMessageFormAsync 的</span><span class="sxs-lookup"><span data-stu-id="eadeb-157">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="eadeb-158">向 `Mailbox` 显示新邮件窗体的对象添加了一个新函数。</span><span class="sxs-lookup"><span data-stu-id="eadeb-158">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="eadeb-159">这是方法的异步版本 `displayNewMessageForm` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-159">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="eadeb-160">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="eadeb-161">DisplayReplyAllFormAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="eadeb-161">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="eadeb-162">向 `Item` 在阅读模式下显示 "全部答复" 窗体的对象添加了一个新函数。</span><span class="sxs-lookup"><span data-stu-id="eadeb-162">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="eadeb-163">这是方法的异步版本 `displayReplyAllForm` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-163">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="eadeb-164">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-164">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="eadeb-165">DisplayReplyFormAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="eadeb-165">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="eadeb-166">向 `Item` 在阅读模式下显示 "答复" 窗体的对象添加了一个新函数。</span><span class="sxs-lookup"><span data-stu-id="eadeb-166">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="eadeb-167">这是方法的异步版本 `displayReplyForm` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-167">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="eadeb-168">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-168">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="eadeb-169">基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="eadeb-169">Event-based activation</span></span>

<span data-ttu-id="eadeb-170">添加了对 Outlook 外接程序中基于事件的激活功能的支持。若要了解详细信息，请参阅 [配置 Outlook 外接程序以进行基于事件的激活](../../../outlook/autolaunch.md) 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-170">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="eadeb-171">LaunchEvent 扩展点</span><span class="sxs-lookup"><span data-stu-id="eadeb-171">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="eadeb-172">`LaunchEvent`向清单添加了扩展点支持。</span><span class="sxs-lookup"><span data-stu-id="eadeb-172">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="eadeb-173">它配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="eadeb-173">It configures event-based activation functionality.</span></span>

<span data-ttu-id="eadeb-174">**中的可用**： Outlook 网页版 (新式的 [请求预览访问](https://aka.ms/OWAPreview)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-174">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="eadeb-175">LaunchEvents 清单元素</span><span class="sxs-lookup"><span data-stu-id="eadeb-175">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="eadeb-176">`LaunchEvents`向清单添加了元素。</span><span class="sxs-lookup"><span data-stu-id="eadeb-176">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="eadeb-177">它支持配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="eadeb-177">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="eadeb-178">**中的可用**： Outlook 网页版 (新式的 [请求预览访问](https://aka.ms/OWAPreview)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-178">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="eadeb-179">运行时清单元素</span><span class="sxs-lookup"><span data-stu-id="eadeb-179">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="eadeb-180">向清单元素添加了 Outlook 支持 `Runtimes` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-180">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="eadeb-181">它引用了基于事件的激活功能所需的 HTML 和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="eadeb-181">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="eadeb-182">**中的可用**： Outlook 网页版 (新式的 [请求预览访问](https://aka.ms/OWAPreview)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-182">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="eadeb-183">获取所有自定义属性</span><span class="sxs-lookup"><span data-stu-id="eadeb-183">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="eadeb-184">CustomProperties。 getAll</span><span class="sxs-lookup"><span data-stu-id="eadeb-184">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true#getall--)

<span data-ttu-id="eadeb-185">向 `CustomProperties` 获取所有自定义属性的对象添加了新函数。</span><span class="sxs-lookup"><span data-stu-id="eadeb-185">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="eadeb-186">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook (网页版) ，Mac 上的 outlook (已连接到 microsoft 365 订阅) ，Outlook 在 Android 上，Outlook 在 iOS 上</span><span class="sxs-lookup"><span data-stu-id="eadeb-186">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="eadeb-187">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="eadeb-187">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="eadeb-188">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="eadeb-188">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="eadeb-189">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="eadeb-189">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="eadeb-190">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (经典) </span><span class="sxs-lookup"><span data-stu-id="eadeb-190">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="eadeb-191">邮件签名</span><span class="sxs-lookup"><span data-stu-id="eadeb-191">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="eadeb-192">SetSignatureAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="eadeb-192">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="eadeb-193">向对象添加了一个新函数 `Body` ，该函数在撰写模式下添加或替换项目正文中的签名。</span><span class="sxs-lookup"><span data-stu-id="eadeb-193">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="eadeb-194">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-194">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="eadeb-195">DisableClientSignatureAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="eadeb-195">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="eadeb-196">添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。</span><span class="sxs-lookup"><span data-stu-id="eadeb-196">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="eadeb-197">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-197">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="eadeb-198">GetComposeTypeAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="eadeb-198">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="eadeb-199">添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。</span><span class="sxs-lookup"><span data-stu-id="eadeb-199">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="eadeb-200">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-200">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="eadeb-201">IsClientSignatureEnabledAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="eadeb-201">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="eadeb-202">添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。</span><span class="sxs-lookup"><span data-stu-id="eadeb-202">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="eadeb-203">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-203">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="eadeb-204">MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="eadeb-204">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="eadeb-205">添加了一个新枚举，该枚举 `ComposeType` 在撰写模式中可用。</span><span class="sxs-lookup"><span data-stu-id="eadeb-205">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="eadeb-206">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式中， [配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)) </span><span class="sxs-lookup"><span data-stu-id="eadeb-206">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="eadeb-207">包含操作的通知邮件</span><span class="sxs-lookup"><span data-stu-id="eadeb-207">Notification messages with actions</span></span>

<span data-ttu-id="eadeb-208">通过此功能，您的外接程序可以在默认 **取消** 操作之外包含具有自定义操作的通知消息。</span><span class="sxs-lookup"><span data-stu-id="eadeb-208">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="eadeb-209">NotificationMessageDetails</span><span class="sxs-lookup"><span data-stu-id="eadeb-209">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="eadeb-210">添加了一个新属性，您可以 `InsightMessage` 使用自定义操作添加通知。</span><span class="sxs-lookup"><span data-stu-id="eadeb-210">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="eadeb-211">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-211">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="eadeb-212">NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="eadeb-212">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="eadeb-213">添加了一个新对象，可在其中为通知定义自定义操作 `InsightMessage` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-213">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="eadeb-214">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-214">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="eadeb-215">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="eadeb-215">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="eadeb-216">添加了新枚举 `ActionType` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-216">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="eadeb-217">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-217">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="eadeb-218">ItemNotificationMessageType InsightMessage</span><span class="sxs-lookup"><span data-stu-id="eadeb-218">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="eadeb-219">向枚举添加了新类型 `InsightMessage` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="eadeb-219">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="eadeb-220">**适用于**： Windows (上的 outlook 连接到 Microsoft 365 订阅) ，outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="eadeb-220">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="eadeb-221">Office 主题</span><span class="sxs-lookup"><span data-stu-id="eadeb-221">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="eadeb-222">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="eadeb-222">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="eadeb-223">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="eadeb-223">Added ability to get Office theme.</span></span>

<span data-ttu-id="eadeb-224">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-224">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="eadeb-225">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="eadeb-225">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="eadeb-226">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="eadeb-226">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="eadeb-227">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-227">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="eadeb-228">会话数据</span><span class="sxs-lookup"><span data-stu-id="eadeb-228">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="eadeb-229">SessionData</span><span class="sxs-lookup"><span data-stu-id="eadeb-229">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="eadeb-230">添加了一个代表项目的会话数据的新对象。</span><span class="sxs-lookup"><span data-stu-id="eadeb-230">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="eadeb-231">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-231">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="eadeb-232">SessionData 的 Office。</span><span class="sxs-lookup"><span data-stu-id="eadeb-232">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="eadeb-233">添加了一个新属性以在撰写模式下管理项目的会话数据。</span><span class="sxs-lookup"><span data-stu-id="eadeb-233">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="eadeb-234">**适用于**： Windows (上的 Outlook 连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="eadeb-234">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="eadeb-235">单一登录 (SSO)</span><span class="sxs-lookup"><span data-stu-id="eadeb-235">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="eadeb-236">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="eadeb-236">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="eadeb-237">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="eadeb-237">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="eadeb-238">**适用于**： Outlook on Windows (连接到 microsoft 365 订阅) ，Mac 上的 outlook (连接到 microsoft 365 订阅) ，outlook 网页版 (新式) ，outlook 网页版 (经典) </span><span class="sxs-lookup"><span data-stu-id="eadeb-238">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="eadeb-239">另请参阅</span><span class="sxs-lookup"><span data-stu-id="eadeb-239">See also</span></span>

- [<span data-ttu-id="eadeb-240">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="eadeb-240">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="eadeb-241">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="eadeb-241">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="eadeb-242">入门</span><span class="sxs-lookup"><span data-stu-id="eadeb-242">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="eadeb-243">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="eadeb-243">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
