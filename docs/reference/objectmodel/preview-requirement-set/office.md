---
title: Office 命名空间 - 预览要求集
description: Office使用邮箱 API 预览要求集Outlook外接程序可用的命名空间成员。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 72e2300dd50ff01e26417efaca92906049358fc0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590881"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="bd1b6-103">Office (邮箱预览要求集) </span><span class="sxs-lookup"><span data-stu-id="bd1b6-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="bd1b6-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd1b6-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="bd1b6-106">Requirements</span></span>

|<span data-ttu-id="bd1b6-107">要求</span><span class="sxs-lookup"><span data-stu-id="bd1b6-107">Requirement</span></span>| <span data-ttu-id="bd1b6-108">值</span><span class="sxs-lookup"><span data-stu-id="bd1b6-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd1b6-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bd1b6-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bd1b6-110">1.1</span><span class="sxs-lookup"><span data-stu-id="bd1b6-110">1.1</span></span>|
|[<span data-ttu-id="bd1b6-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bd1b6-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bd1b6-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="bd1b6-113">属性</span><span class="sxs-lookup"><span data-stu-id="bd1b6-113">Properties</span></span>

| <span data-ttu-id="bd1b6-114">属性</span><span class="sxs-lookup"><span data-stu-id="bd1b6-114">Property</span></span> | <span data-ttu-id="bd1b6-115">模式</span><span class="sxs-lookup"><span data-stu-id="bd1b6-115">Modes</span></span> | <span data-ttu-id="bd1b6-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-116">Return type</span></span> | <span data-ttu-id="bd1b6-117">最小值</span><span class="sxs-lookup"><span data-stu-id="bd1b6-117">Minimum</span></span><br><span data-ttu-id="bd1b6-118">要求集</span><span class="sxs-lookup"><span data-stu-id="bd1b6-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bd1b6-119">context</span><span class="sxs-lookup"><span data-stu-id="bd1b6-119">context</span></span>](office.context.md) | <span data-ttu-id="bd1b6-120">撰写</span><span class="sxs-lookup"><span data-stu-id="bd1b6-120">Compose</span></span><br><span data-ttu-id="bd1b6-121">阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-121">Read</span></span> | [<span data-ttu-id="bd1b6-122">Context</span><span class="sxs-lookup"><span data-stu-id="bd1b6-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="bd1b6-123">1.1</span><span class="sxs-lookup"><span data-stu-id="bd1b6-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="bd1b6-124">枚举</span><span class="sxs-lookup"><span data-stu-id="bd1b6-124">Enumerations</span></span>

| <span data-ttu-id="bd1b6-125">枚举</span><span class="sxs-lookup"><span data-stu-id="bd1b6-125">Enumeration</span></span> | <span data-ttu-id="bd1b6-126">模式</span><span class="sxs-lookup"><span data-stu-id="bd1b6-126">Modes</span></span> | <span data-ttu-id="bd1b6-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-127">Return type</span></span> | <span data-ttu-id="bd1b6-128">最小值</span><span class="sxs-lookup"><span data-stu-id="bd1b6-128">Minimum</span></span><br><span data-ttu-id="bd1b6-129">要求集</span><span class="sxs-lookup"><span data-stu-id="bd1b6-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bd1b6-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="bd1b6-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="bd1b6-131">撰写</span><span class="sxs-lookup"><span data-stu-id="bd1b6-131">Compose</span></span><br><span data-ttu-id="bd1b6-132">阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-132">Read</span></span> | <span data-ttu-id="bd1b6-133">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-133">String</span></span> | [<span data-ttu-id="bd1b6-134">1.1</span><span class="sxs-lookup"><span data-stu-id="bd1b6-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bd1b6-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="bd1b6-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="bd1b6-136">撰写</span><span class="sxs-lookup"><span data-stu-id="bd1b6-136">Compose</span></span><br><span data-ttu-id="bd1b6-137">阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-137">Read</span></span> | <span data-ttu-id="bd1b6-138">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-138">String</span></span> | [<span data-ttu-id="bd1b6-139">1.1</span><span class="sxs-lookup"><span data-stu-id="bd1b6-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bd1b6-140">EventType</span><span class="sxs-lookup"><span data-stu-id="bd1b6-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="bd1b6-141">撰写</span><span class="sxs-lookup"><span data-stu-id="bd1b6-141">Compose</span></span><br><span data-ttu-id="bd1b6-142">阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-142">Read</span></span> | <span data-ttu-id="bd1b6-143">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-143">String</span></span> | [<span data-ttu-id="bd1b6-144">1.5</span><span class="sxs-lookup"><span data-stu-id="bd1b6-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="bd1b6-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="bd1b6-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="bd1b6-146">撰写</span><span class="sxs-lookup"><span data-stu-id="bd1b6-146">Compose</span></span><br><span data-ttu-id="bd1b6-147">阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-147">Read</span></span> | <span data-ttu-id="bd1b6-148">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-148">String</span></span> | [<span data-ttu-id="bd1b6-149">1.1</span><span class="sxs-lookup"><span data-stu-id="bd1b6-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="bd1b6-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="bd1b6-150">Namespaces</span></span>

<span data-ttu-id="bd1b6-151">[MailboxEnums：](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true)包括许多Outlook枚举，例如、 `ItemType` `EntityType` 和 `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="bd1b6-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="bd1b6-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="bd1b6-153">AsyncResultStatus：String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="bd1b6-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="bd1b6-155">类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-155">Type</span></span>

*   <span data-ttu-id="bd1b6-156">String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bd1b6-157">属性</span><span class="sxs-lookup"><span data-stu-id="bd1b6-157">Properties</span></span>

|<span data-ttu-id="bd1b6-158">名称</span><span class="sxs-lookup"><span data-stu-id="bd1b6-158">Name</span></span>| <span data-ttu-id="bd1b6-159">类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-159">Type</span></span>| <span data-ttu-id="bd1b6-160">描述</span><span class="sxs-lookup"><span data-stu-id="bd1b6-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="bd1b6-161">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-161">String</span></span>|<span data-ttu-id="bd1b6-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="bd1b6-163">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-163">String</span></span>|<span data-ttu-id="bd1b6-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd1b6-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="bd1b6-165">Requirements</span></span>

|<span data-ttu-id="bd1b6-166">要求</span><span class="sxs-lookup"><span data-stu-id="bd1b6-166">Requirement</span></span>| <span data-ttu-id="bd1b6-167">值</span><span class="sxs-lookup"><span data-stu-id="bd1b6-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd1b6-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bd1b6-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bd1b6-169">1.1</span><span class="sxs-lookup"><span data-stu-id="bd1b6-169">1.1</span></span>|
|[<span data-ttu-id="bd1b6-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bd1b6-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bd1b6-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="bd1b6-172">CoercionType：String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-172">CoercionType: String</span></span>

<span data-ttu-id="bd1b6-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bd1b6-174">类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-174">Type</span></span>

*   <span data-ttu-id="bd1b6-175">String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bd1b6-176">属性</span><span class="sxs-lookup"><span data-stu-id="bd1b6-176">Properties</span></span>

|<span data-ttu-id="bd1b6-177">名称</span><span class="sxs-lookup"><span data-stu-id="bd1b6-177">Name</span></span>| <span data-ttu-id="bd1b6-178">类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-178">Type</span></span>| <span data-ttu-id="bd1b6-179">描述</span><span class="sxs-lookup"><span data-stu-id="bd1b6-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="bd1b6-180">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-180">String</span></span>|<span data-ttu-id="bd1b6-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="bd1b6-182">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-182">String</span></span>|<span data-ttu-id="bd1b6-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd1b6-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="bd1b6-184">Requirements</span></span>

|<span data-ttu-id="bd1b6-185">要求</span><span class="sxs-lookup"><span data-stu-id="bd1b6-185">Requirement</span></span>| <span data-ttu-id="bd1b6-186">值</span><span class="sxs-lookup"><span data-stu-id="bd1b6-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd1b6-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bd1b6-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bd1b6-188">1.1</span><span class="sxs-lookup"><span data-stu-id="bd1b6-188">1.1</span></span>|
|[<span data-ttu-id="bd1b6-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bd1b6-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bd1b6-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="bd1b6-191">EventType：String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-191">EventType: String</span></span>

<span data-ttu-id="bd1b6-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="bd1b6-193">类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-193">Type</span></span>

*   <span data-ttu-id="bd1b6-194">String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bd1b6-195">属性</span><span class="sxs-lookup"><span data-stu-id="bd1b6-195">Properties</span></span>

| <span data-ttu-id="bd1b6-196">名称</span><span class="sxs-lookup"><span data-stu-id="bd1b6-196">Name</span></span> | <span data-ttu-id="bd1b6-197">类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-197">Type</span></span> | <span data-ttu-id="bd1b6-198">描述</span><span class="sxs-lookup"><span data-stu-id="bd1b6-198">Description</span></span> | <span data-ttu-id="bd1b6-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="bd1b6-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="bd1b6-200">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-200">String</span></span> | <span data-ttu-id="bd1b6-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="bd1b6-202">1.7</span><span class="sxs-lookup"><span data-stu-id="bd1b6-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="bd1b6-203">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-203">String</span></span> | <span data-ttu-id="bd1b6-204">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="bd1b6-205">1.8</span><span class="sxs-lookup"><span data-stu-id="bd1b6-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="bd1b6-206">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-206">String</span></span> | <span data-ttu-id="bd1b6-207">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="bd1b6-208">1.8</span><span class="sxs-lookup"><span data-stu-id="bd1b6-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="bd1b6-209">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-209">String</span></span> | <span data-ttu-id="bd1b6-210">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="bd1b6-211">1.5</span><span class="sxs-lookup"><span data-stu-id="bd1b6-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="bd1b6-212">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-212">String</span></span> | <span data-ttu-id="bd1b6-213">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="bd1b6-214">预览</span><span class="sxs-lookup"><span data-stu-id="bd1b6-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="bd1b6-215">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-215">String</span></span> | <span data-ttu-id="bd1b6-216">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="bd1b6-217">1.7</span><span class="sxs-lookup"><span data-stu-id="bd1b6-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="bd1b6-218">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-218">String</span></span> | <span data-ttu-id="bd1b6-219">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="bd1b6-220">1.7</span><span class="sxs-lookup"><span data-stu-id="bd1b6-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bd1b6-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="bd1b6-221">Requirements</span></span>

|<span data-ttu-id="bd1b6-222">要求</span><span class="sxs-lookup"><span data-stu-id="bd1b6-222">Requirement</span></span>| <span data-ttu-id="bd1b6-223">值</span><span class="sxs-lookup"><span data-stu-id="bd1b6-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd1b6-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bd1b6-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bd1b6-225">1.5</span><span class="sxs-lookup"><span data-stu-id="bd1b6-225">1.5</span></span> |
|[<span data-ttu-id="bd1b6-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bd1b6-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bd1b6-227">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="bd1b6-228">SourceProperty：String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-228">SourceProperty: String</span></span>

<span data-ttu-id="bd1b6-229">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bd1b6-230">类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-230">Type</span></span>

*   <span data-ttu-id="bd1b6-231">String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bd1b6-232">属性</span><span class="sxs-lookup"><span data-stu-id="bd1b6-232">Properties</span></span>

|<span data-ttu-id="bd1b6-233">名称</span><span class="sxs-lookup"><span data-stu-id="bd1b6-233">Name</span></span>| <span data-ttu-id="bd1b6-234">类型</span><span class="sxs-lookup"><span data-stu-id="bd1b6-234">Type</span></span>| <span data-ttu-id="bd1b6-235">描述</span><span class="sxs-lookup"><span data-stu-id="bd1b6-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="bd1b6-236">字符串</span><span class="sxs-lookup"><span data-stu-id="bd1b6-236">String</span></span>|<span data-ttu-id="bd1b6-237">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="bd1b6-238">String</span><span class="sxs-lookup"><span data-stu-id="bd1b6-238">String</span></span>|<span data-ttu-id="bd1b6-239">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="bd1b6-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd1b6-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="bd1b6-240">Requirements</span></span>

|<span data-ttu-id="bd1b6-241">要求</span><span class="sxs-lookup"><span data-stu-id="bd1b6-241">Requirement</span></span>| <span data-ttu-id="bd1b6-242">值</span><span class="sxs-lookup"><span data-stu-id="bd1b6-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd1b6-243">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bd1b6-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bd1b6-244">1.1</span><span class="sxs-lookup"><span data-stu-id="bd1b6-244">1.1</span></span>|
|[<span data-ttu-id="bd1b6-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bd1b6-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bd1b6-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bd1b6-246">Compose or Read</span></span>|
