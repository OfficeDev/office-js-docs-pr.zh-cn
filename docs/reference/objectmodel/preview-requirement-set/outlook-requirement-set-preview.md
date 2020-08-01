---
title: Outlook 外接程序 API 预览要求集
description: 当前在 Outlook 外接程序的预览中的功能和 Api。
ms.date: 07/22/2020
localization_priority: Normal
ms.openlocfilehash: 9a2ca9a71498c954d9864f2b967b5fb240f8c506
ms.sourcegitcommit: 7d5407d3900d2ad1feae79a4bc038afe50568be0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2020
ms.locfileid: "46530476"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="f3119-103">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="f3119-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="f3119-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="f3119-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f3119-105">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="f3119-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="f3119-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="f3119-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="f3119-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="f3119-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="f3119-108">您可以通过[在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)来预览 Web 上 Outlook 中的功能。</span><span class="sxs-lookup"><span data-stu-id="f3119-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="f3119-109">此页面上的 "配置预览访问权限" 对适用的功能进行了说明。</span><span class="sxs-lookup"><span data-stu-id="f3119-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="f3119-110">对于其他功能，你可以通过填写和提交[此表单](https://aka.ms/OWAPreview)，使用 Microsoft 365 帐户请求对网站上的 Outlook 的预览位的访问权限。</span><span class="sxs-lookup"><span data-stu-id="f3119-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="f3119-111">这些功能上记录了 "请求预览访问"。</span><span class="sxs-lookup"><span data-stu-id="f3119-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="f3119-112">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="f3119-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="f3119-113">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="f3119-113">Features in preview</span></span>

<span data-ttu-id="f3119-114">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="f3119-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="f3119-115">其他日历属性</span><span class="sxs-lookup"><span data-stu-id="f3119-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="f3119-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="f3119-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="f3119-117">在撰写模式下添加了一个代表约会全天事件属性的新对象。</span><span class="sxs-lookup"><span data-stu-id="f3119-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="f3119-118">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-118">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="f3119-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="f3119-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="f3119-120">添加了一个新对象，该对象表示在撰写模式下约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="f3119-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="f3119-121">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-121">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="f3119-122">IsAllDayEvent 的 Office。</span><span class="sxs-lookup"><span data-stu-id="f3119-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="f3119-123">添加了一个新属性，该属性表示约会是否为全天事件。</span><span class="sxs-lookup"><span data-stu-id="f3119-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="f3119-124">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-124">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="f3119-125">"Context"。项目敏感度</span><span class="sxs-lookup"><span data-stu-id="f3119-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="f3119-126">添加了一个表示约会敏感度的新属性。</span><span class="sxs-lookup"><span data-stu-id="f3119-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="f3119-127">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-127">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="f3119-128">MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="f3119-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="f3119-129">添加了一个 `AppointmentSensitivityType` 代表约会上可用的敏感度选项的新枚举。</span><span class="sxs-lookup"><span data-stu-id="f3119-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="f3119-130">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-130">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="f3119-131">发送时追加</span><span class="sxs-lookup"><span data-stu-id="f3119-131">Append on send</span></span>

<span data-ttu-id="f3119-132">若要了解如何使用 "发送时追加" 功能，请参阅在[Outlook 加载项中实施 "在发送时实现附加](../../../outlook/append-on-send.md)"。</span><span class="sxs-lookup"><span data-stu-id="f3119-132">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="f3119-133">AppendOnSendAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="f3119-133">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="f3119-134">向对象添加了一个新函数 `Body` ，该函数在撰写模式下将数据追加到项正文的末尾。</span><span class="sxs-lookup"><span data-stu-id="f3119-134">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="f3119-135">**适用于**： Windows 上的 outlook （连接到 Microsoft 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="f3119-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="f3119-136">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="f3119-136">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="f3119-137">向清单添加了一个新元素，其中 `AppendOnSend` 扩展权限必须包含在扩展权限的集合中。</span><span class="sxs-lookup"><span data-stu-id="f3119-137">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="f3119-138">**适用于**： Windows 上的 outlook （连接到 Microsoft 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="f3119-138">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="f3119-139">Api 的异步版本 `display`</span><span class="sxs-lookup"><span data-stu-id="f3119-139">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="f3119-140">DisplayAppointmentFormAsync 的</span><span class="sxs-lookup"><span data-stu-id="f3119-140">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="f3119-141">向显示现有约会的对象添加了新函数 `Mailbox` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-141">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="f3119-142">这是方法的异步版本 `displayAppointmentForm` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-142">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="f3119-143">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="f3119-144">DisplayMessageFormAsync 的</span><span class="sxs-lookup"><span data-stu-id="f3119-144">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="f3119-145">向显示现有邮件的对象添加了新函数 `Mailbox` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-145">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="f3119-146">这是方法的异步版本 `displayMessageForm` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-146">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="f3119-147">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-147">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="f3119-148">DisplayNewAppointmentFormAsync 的</span><span class="sxs-lookup"><span data-stu-id="f3119-148">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="f3119-149">向 `Mailbox` 显示新约会窗体的对象添加了一个新函数。</span><span class="sxs-lookup"><span data-stu-id="f3119-149">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="f3119-150">这是方法的异步版本 `displayNewAppointmentForm` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-150">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="f3119-151">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-151">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="f3119-152">DisplayNewMessageFormAsync 的</span><span class="sxs-lookup"><span data-stu-id="f3119-152">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="f3119-153">向 `Mailbox` 显示新邮件窗体的对象添加了一个新函数。</span><span class="sxs-lookup"><span data-stu-id="f3119-153">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="f3119-154">这是方法的异步版本 `displayNewMessageForm` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-154">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="f3119-155">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-155">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="f3119-156">DisplayReplyAllFormAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="f3119-156">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="f3119-157">向 `Item` 在阅读模式下显示 "全部答复" 窗体的对象添加了一个新函数。</span><span class="sxs-lookup"><span data-stu-id="f3119-157">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="f3119-158">这是方法的异步版本 `displayReplyAllForm` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-158">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="f3119-159">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-159">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="f3119-160">DisplayReplyFormAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="f3119-160">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="f3119-161">向 `Item` 在阅读模式下显示 "答复" 窗体的对象添加了一个新函数。</span><span class="sxs-lookup"><span data-stu-id="f3119-161">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="f3119-162">这是方法的异步版本 `displayReplyForm` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-162">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="f3119-163">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="f3119-164">基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="f3119-164">Event-based activation</span></span>

<span data-ttu-id="f3119-165">添加了对 Outlook 外接程序中基于事件的激活功能的支持。若要了解详细信息，请参阅[配置 Outlook 外接程序以进行基于事件的激活](../../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="f3119-165">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="f3119-166">LaunchEvent 扩展点</span><span class="sxs-lookup"><span data-stu-id="f3119-166">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="f3119-167">`LaunchEvent`向清单添加了扩展点支持。</span><span class="sxs-lookup"><span data-stu-id="f3119-167">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="f3119-168">它配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="f3119-168">It configures event-based activation functionality.</span></span>

<span data-ttu-id="f3119-169">**适用于**： Outlook 网页版（新式，[请求预览访问](https://aka.ms/OWAPreview)）</span><span class="sxs-lookup"><span data-stu-id="f3119-169">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="f3119-170">LaunchEvents 清单元素</span><span class="sxs-lookup"><span data-stu-id="f3119-170">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="f3119-171">`LaunchEvents`向清单添加了元素。</span><span class="sxs-lookup"><span data-stu-id="f3119-171">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="f3119-172">它支持配置基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="f3119-172">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="f3119-173">**适用于**： Outlook 网页版（新式，[请求预览访问](https://aka.ms/OWAPreview)）</span><span class="sxs-lookup"><span data-stu-id="f3119-173">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="f3119-174">运行时清单元素</span><span class="sxs-lookup"><span data-stu-id="f3119-174">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="f3119-175">向清单元素添加了 Outlook 支持 `Runtimes` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-175">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="f3119-176">它引用了基于事件的激活功能所需的 HTML 和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="f3119-176">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="f3119-177">**适用于**： Outlook 网页版（新式，[请求预览访问](https://aka.ms/OWAPreview)）</span><span class="sxs-lookup"><span data-stu-id="f3119-177">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="f3119-178">获取所有自定义属性</span><span class="sxs-lookup"><span data-stu-id="f3119-178">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="f3119-179">CustomProperties。 getAll</span><span class="sxs-lookup"><span data-stu-id="f3119-179">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

<span data-ttu-id="f3119-180">向 `CustomProperties` 获取所有自定义属性的对象添加了新函数。</span><span class="sxs-lookup"><span data-stu-id="f3119-180">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="f3119-181">**提供**时间： Windows 上的 outlook （连接到 Microsoft 365 订阅）、web 上的 outlook （新式）、Mac 上的 outlook （已连接到 Microsoft 365 订阅）、Outlook on Android、在 iOS 上的 outlook</span><span class="sxs-lookup"><span data-stu-id="f3119-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="f3119-182">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="f3119-182">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="f3119-183">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="f3119-183">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="f3119-184">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="f3119-184">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="f3119-185">**适用于**： Windows 上的 outlook （连接到 Microsoft 365 订阅），outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="f3119-185">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="f3119-186">邮件签名</span><span class="sxs-lookup"><span data-stu-id="f3119-186">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="f3119-187">SetSignatureAsync 的 "."</span><span class="sxs-lookup"><span data-stu-id="f3119-187">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="f3119-188">向对象添加了一个新函数 `Body` ，该函数在撰写模式下添加或替换项目正文中的签名。</span><span class="sxs-lookup"><span data-stu-id="f3119-188">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="f3119-189">**适用于**： Windows 上的 outlook （连接到 Microsoft 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="f3119-189">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="f3119-190">DisableClientSignatureAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="f3119-190">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="f3119-191">添加了一个新函数，用于在撰写模式下禁用发送邮箱的客户端签名。</span><span class="sxs-lookup"><span data-stu-id="f3119-191">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="f3119-192">**适用于**： Windows 上的 outlook （连接到 Microsoft 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="f3119-192">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="f3119-193">GetComposeTypeAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="f3119-193">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="f3119-194">添加了一个新函数，用于在撰写模式下获取邮件的撰写类型。</span><span class="sxs-lookup"><span data-stu-id="f3119-194">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="f3119-195">**适用于**： Windows 上的 outlook （连接到 Microsoft 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="f3119-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="f3119-196">IsClientSignatureEnabledAsync 的 Office。</span><span class="sxs-lookup"><span data-stu-id="f3119-196">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="f3119-197">添加了一个新函数，用于检查在撰写模式下是否在项目上启用了客户端签名。</span><span class="sxs-lookup"><span data-stu-id="f3119-197">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="f3119-198">**适用于**： Windows 上的 outlook （连接到 Microsoft 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="f3119-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="f3119-199">MailboxEnums. ComposeType</span><span class="sxs-lookup"><span data-stu-id="f3119-199">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="f3119-200">添加了一个新枚举，该枚举 `ComposeType` 在撰写模式中可用。</span><span class="sxs-lookup"><span data-stu-id="f3119-200">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="f3119-201">**适用于**： Windows 上的 outlook （连接到 Microsoft 365 订阅），outlook 网页版（新式，[配置预览访问](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)）</span><span class="sxs-lookup"><span data-stu-id="f3119-201">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="f3119-202">包含操作的通知邮件</span><span class="sxs-lookup"><span data-stu-id="f3119-202">Notification messages with actions</span></span>

<span data-ttu-id="f3119-203">通过此功能，您的外接程序可以在默认**取消**操作之外包含具有自定义操作的通知消息。</span><span class="sxs-lookup"><span data-stu-id="f3119-203">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="f3119-204">NotificationMessageDetails</span><span class="sxs-lookup"><span data-stu-id="f3119-204">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="f3119-205">添加了一个新属性，您可以 `InsightMessage` 使用自定义操作添加通知。</span><span class="sxs-lookup"><span data-stu-id="f3119-205">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="f3119-206">**提供**时间： Windows 上的 outlook （连接到 Microsoft 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="f3119-206">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="f3119-207">NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="f3119-207">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="f3119-208">添加了一个新对象，可在其中为通知定义自定义操作 `InsightMessage` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-208">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="f3119-209">**提供**时间： Windows 上的 outlook （连接到 Microsoft 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="f3119-209">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="f3119-210">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="f3119-210">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="f3119-211">添加了新枚举 `ActionType` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-211">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="f3119-212">**提供**时间： Windows 上的 outlook （连接到 Microsoft 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="f3119-212">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="f3119-213">ItemNotificationMessageType InsightMessage</span><span class="sxs-lookup"><span data-stu-id="f3119-213">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="f3119-214">向枚举添加了新类型 `InsightMessage` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="f3119-214">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="f3119-215">**提供**时间： Windows 上的 outlook （连接到 Microsoft 365 订阅）、outlook 网页版（新式）</span><span class="sxs-lookup"><span data-stu-id="f3119-215">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="f3119-216">Office 主题</span><span class="sxs-lookup"><span data-stu-id="f3119-216">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="f3119-217">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="f3119-217">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="f3119-218">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="f3119-218">Added ability to get Office theme.</span></span>

<span data-ttu-id="f3119-219">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-219">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="f3119-220">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="f3119-220">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f3119-221">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="f3119-221">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="f3119-222">**适用于**： Windows 上的 Outlook （连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="f3119-222">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="f3119-223">单一登录 (SSO)</span><span class="sxs-lookup"><span data-stu-id="f3119-223">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="f3119-224">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="f3119-224">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="f3119-225">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="f3119-225">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="f3119-226">**提供**时间： Windows 上的 outlook （连接到 microsoft 365 订阅）、365 Mac 上的 outlook （新式）（新式）、outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="f3119-226">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="f3119-227">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f3119-227">See also</span></span>

- [<span data-ttu-id="f3119-228">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="f3119-228">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="f3119-229">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="f3119-229">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="f3119-230">入门</span><span class="sxs-lookup"><span data-stu-id="f3119-230">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="f3119-231">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="f3119-231">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
