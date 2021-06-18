---
title: Outlook外接程序 API 预览要求集
description: 当前处于预览阶段的功能和 API Outlook外接程序。
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: f9d8afc2b4347a8fb13f8ab98a163fb63968123f
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007760"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="7f271-103">Outlook外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="7f271-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="7f271-104">Outlook JavaScript API 的 Office 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="7f271-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7f271-105">本文档适用于 **预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="7f271-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="7f271-106">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="7f271-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="7f271-107">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="7f271-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="7f271-108">你或许能够通过在 Outlook 网页版 租户上配置定向发布来预览 Microsoft 365[功能](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="7f271-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="7f271-109">此页面中会针对适用的功能说明"配置预览访问"。</span><span class="sxs-lookup"><span data-stu-id="7f271-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="7f271-110">对于其他功能，您可能能够通过完成和提交此表单Outlook 网页版请求访问 Microsoft 365 预览[位](https://aka.ms/OWAPreview)。</span><span class="sxs-lookup"><span data-stu-id="7f271-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="7f271-111">这些功能中会指出"请求预览访问"。</span><span class="sxs-lookup"><span data-stu-id="7f271-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="7f271-112">预览要求集包含要求集 [1.10 的所有功能](../requirement-set-1.10/outlook-requirement-set-1.10.md)。</span><span class="sxs-lookup"><span data-stu-id="7f271-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="7f271-113">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="7f271-113">Features in preview</span></span>

<span data-ttu-id="7f271-114">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="7f271-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="7f271-115">对受信息权限管理中心 IRM 保护的项目 (加载项) </span><span class="sxs-lookup"><span data-stu-id="7f271-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="7f271-116">现在可以在受 IRM 保护的项目上激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="7f271-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="7f271-117">若要启用此功能，租户管理员需要在租户中设置"允许以编程方式访问"自定义策略选项， `OBJMODEL` 以启用Office。 </span><span class="sxs-lookup"><span data-stu-id="7f271-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="7f271-118">有关详细信息 [，请参阅使用](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 权限和说明。</span><span class="sxs-lookup"><span data-stu-id="7f271-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="7f271-119">**提供位置**：Outlook Windows版本 13229.10000 (连接到 Microsoft 365 订阅) </span><span class="sxs-lookup"><span data-stu-id="7f271-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="7f271-120">其他日历属性</span><span class="sxs-lookup"><span data-stu-id="7f271-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="7f271-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="7f271-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="7f271-122">添加了一个新对象，该对象代表撰写模式下约会的全天事件属性。</span><span class="sxs-lookup"><span data-stu-id="7f271-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="7f271-123">**在**：Outlook Windows (订阅Microsoft 365上) </span><span class="sxs-lookup"><span data-stu-id="7f271-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="7f271-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="7f271-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="7f271-125">添加了一个新对象，该对象表示撰写模式下约会的敏感度。</span><span class="sxs-lookup"><span data-stu-id="7f271-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="7f271-126">**在**：Outlook Windows (订阅Microsoft 365上) </span><span class="sxs-lookup"><span data-stu-id="7f271-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="7f271-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="7f271-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="7f271-128">添加了一个新属性，它表示约会是全天事件。</span><span class="sxs-lookup"><span data-stu-id="7f271-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="7f271-129">**在**：Outlook Windows (订阅Microsoft 365上) </span><span class="sxs-lookup"><span data-stu-id="7f271-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="7f271-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="7f271-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="7f271-131">新增了一个表示约会敏感度的属性。</span><span class="sxs-lookup"><span data-stu-id="7f271-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="7f271-132">**在**：Outlook Windows (订阅Microsoft 365上) </span><span class="sxs-lookup"><span data-stu-id="7f271-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="7f271-133">Office。MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="7f271-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="7f271-134">新增了表示约会 `AppointmentSensitivityType` 可用的敏感度选项的枚举。</span><span class="sxs-lookup"><span data-stu-id="7f271-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="7f271-135">**在**：Outlook Windows (订阅Microsoft 365上) </span><span class="sxs-lookup"><span data-stu-id="7f271-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="7f271-136">基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="7f271-136">Event-based activation</span></span>

<span data-ttu-id="7f271-137">此功能在要求集 [1.10 中发布](../requirement-set-1.10/outlook-requirement-set-1.10.md)。</span><span class="sxs-lookup"><span data-stu-id="7f271-137">This feature was released in [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="7f271-138">不过，其他事件现在预览版中可用。</span><span class="sxs-lookup"><span data-stu-id="7f271-138">However, additional events are now available in preview.</span></span> <span data-ttu-id="7f271-139">若要了解详情，请参阅 [支持的事件](../../../outlook/autolaunch.md#supported-events)。</span><span class="sxs-lookup"><span data-stu-id="7f271-139">To learn more, refer to [Supported events](../../../outlook/autolaunch.md#supported-events).</span></span>

<span data-ttu-id="7f271-140">**适用于**：Outlook Windows (连接到 Microsoft 365 订阅) ，Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="7f271-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="7f271-141">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="7f271-141">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="7f271-142">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="7f271-142">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="7f271-143">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="7f271-143">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="7f271-144">**适用于**：Outlook Windows (连接到 Microsoft 365 订阅) ，Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="7f271-144">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="7f271-145">Office 主题</span><span class="sxs-lookup"><span data-stu-id="7f271-145">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="7f271-146">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="7f271-146">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="7f271-147">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="7f271-147">Added ability to get Office theme.</span></span>

<span data-ttu-id="7f271-148">**在**：Outlook Windows (订阅Microsoft 365上) </span><span class="sxs-lookup"><span data-stu-id="7f271-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="7f271-149">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="7f271-149">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="7f271-150">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="7f271-150">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="7f271-151">**在**：Outlook Windows (订阅Microsoft 365上) </span><span class="sxs-lookup"><span data-stu-id="7f271-151">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="7f271-152">会话数据</span><span class="sxs-lookup"><span data-stu-id="7f271-152">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="7f271-153">Office。SessionData</span><span class="sxs-lookup"><span data-stu-id="7f271-153">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="7f271-154">添加了一个新对象，该对象表示项目的会话数据。</span><span class="sxs-lookup"><span data-stu-id="7f271-154">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="7f271-155">**适用于**：Outlook Windows (连接到 Microsoft 365 订阅) ，Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="7f271-155">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="7f271-156">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="7f271-156">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="7f271-157">添加了一个新属性，用于管理撰写模式下项目的会话数据。</span><span class="sxs-lookup"><span data-stu-id="7f271-157">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="7f271-158">**适用于**：Outlook Windows (连接到 Microsoft 365 订阅) ，Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="7f271-158">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="shared-mailboxes"></a><span data-ttu-id="7f271-159">共享邮箱</span><span class="sxs-lookup"><span data-stu-id="7f271-159">Shared mailboxes</span></span>

<span data-ttu-id="7f271-160">共享文件夹的功能支持 (，即委派访问权限) 要求集 [1.8 中发布](../requirement-set-1.8/outlook-requirement-set-1.8.md)。</span><span class="sxs-lookup"><span data-stu-id="7f271-160">Feature support for shared folders (that is, delegate access) was released in [requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="7f271-161">但是，现在预览版中提供了对共享邮箱的支持。</span><span class="sxs-lookup"><span data-stu-id="7f271-161">However, support for shared mailboxes is now available in preview.</span></span> <span data-ttu-id="7f271-162">若要了解详情，请参阅启用 [共享文件夹和共享邮箱方案](../../../outlook/delegate-access.md)。</span><span class="sxs-lookup"><span data-stu-id="7f271-162">To learn more, refer to [Enable shared folders and shared mailbox scenarios](../../../outlook/delegate-access.md).</span></span>

<span data-ttu-id="7f271-163">**适用于**：Outlook Windows (连接到 Microsoft 365 订阅) ，Outlook 网页版 (新式) </span><span class="sxs-lookup"><span data-stu-id="7f271-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="7f271-164">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7f271-164">See also</span></span>

- [<span data-ttu-id="7f271-165">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="7f271-165">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="7f271-166">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="7f271-166">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="7f271-167">入门</span><span class="sxs-lookup"><span data-stu-id="7f271-167">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="7f271-168">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="7f271-168">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
