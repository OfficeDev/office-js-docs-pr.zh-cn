---
title: Office命名空间 - 要求集 1.8
description: Office邮箱 API 要求集 1.8 Outlook外接程序可用的命名空间成员。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 00e236bed7e00159be8c94f727ca64ccaecd07b0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590524"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="4d06e-103">Office (邮箱要求集 1.8) </span><span class="sxs-lookup"><span data-stu-id="4d06e-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="4d06e-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="4d06e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4d06e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4d06e-106">Requirements</span></span>

|<span data-ttu-id="4d06e-107">要求</span><span class="sxs-lookup"><span data-stu-id="4d06e-107">Requirement</span></span>| <span data-ttu-id="4d06e-108">值</span><span class="sxs-lookup"><span data-stu-id="4d06e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d06e-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4d06e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4d06e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="4d06e-110">1.1</span></span>|
|[<span data-ttu-id="4d06e-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4d06e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4d06e-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="4d06e-113">属性</span><span class="sxs-lookup"><span data-stu-id="4d06e-113">Properties</span></span>

| <span data-ttu-id="4d06e-114">属性</span><span class="sxs-lookup"><span data-stu-id="4d06e-114">Property</span></span> | <span data-ttu-id="4d06e-115">模式</span><span class="sxs-lookup"><span data-stu-id="4d06e-115">Modes</span></span> | <span data-ttu-id="4d06e-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-116">Return type</span></span> | <span data-ttu-id="4d06e-117">最小值</span><span class="sxs-lookup"><span data-stu-id="4d06e-117">Minimum</span></span><br><span data-ttu-id="4d06e-118">要求集</span><span class="sxs-lookup"><span data-stu-id="4d06e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4d06e-119">context</span><span class="sxs-lookup"><span data-stu-id="4d06e-119">context</span></span>](office.context.md) | <span data-ttu-id="4d06e-120">撰写</span><span class="sxs-lookup"><span data-stu-id="4d06e-120">Compose</span></span><br><span data-ttu-id="4d06e-121">阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-121">Read</span></span> | [<span data-ttu-id="4d06e-122">Context</span><span class="sxs-lookup"><span data-stu-id="4d06e-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="4d06e-123">1.1</span><span class="sxs-lookup"><span data-stu-id="4d06e-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="4d06e-124">枚举</span><span class="sxs-lookup"><span data-stu-id="4d06e-124">Enumerations</span></span>

| <span data-ttu-id="4d06e-125">枚举</span><span class="sxs-lookup"><span data-stu-id="4d06e-125">Enumeration</span></span> | <span data-ttu-id="4d06e-126">模式</span><span class="sxs-lookup"><span data-stu-id="4d06e-126">Modes</span></span> | <span data-ttu-id="4d06e-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-127">Return type</span></span> | <span data-ttu-id="4d06e-128">最小值</span><span class="sxs-lookup"><span data-stu-id="4d06e-128">Minimum</span></span><br><span data-ttu-id="4d06e-129">要求集</span><span class="sxs-lookup"><span data-stu-id="4d06e-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4d06e-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4d06e-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4d06e-131">撰写</span><span class="sxs-lookup"><span data-stu-id="4d06e-131">Compose</span></span><br><span data-ttu-id="4d06e-132">阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-132">Read</span></span> | <span data-ttu-id="4d06e-133">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-133">String</span></span> | [<span data-ttu-id="4d06e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="4d06e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4d06e-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4d06e-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4d06e-136">撰写</span><span class="sxs-lookup"><span data-stu-id="4d06e-136">Compose</span></span><br><span data-ttu-id="4d06e-137">阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-137">Read</span></span> | <span data-ttu-id="4d06e-138">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-138">String</span></span> | [<span data-ttu-id="4d06e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="4d06e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4d06e-140">EventType</span><span class="sxs-lookup"><span data-stu-id="4d06e-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4d06e-141">撰写</span><span class="sxs-lookup"><span data-stu-id="4d06e-141">Compose</span></span><br><span data-ttu-id="4d06e-142">阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-142">Read</span></span> | <span data-ttu-id="4d06e-143">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-143">String</span></span> | [<span data-ttu-id="4d06e-144">1.5</span><span class="sxs-lookup"><span data-stu-id="4d06e-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="4d06e-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4d06e-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4d06e-146">撰写</span><span class="sxs-lookup"><span data-stu-id="4d06e-146">Compose</span></span><br><span data-ttu-id="4d06e-147">阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-147">Read</span></span> | <span data-ttu-id="4d06e-148">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-148">String</span></span> | [<span data-ttu-id="4d06e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="4d06e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="4d06e-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="4d06e-150">Namespaces</span></span>

<span data-ttu-id="4d06e-151">[MailboxEnums：](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true)包括许多Outlook枚举，例如、 `ItemType` `EntityType` 和 `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="4d06e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="4d06e-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="4d06e-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4d06e-153">AsyncResultStatus：String</span><span class="sxs-lookup"><span data-stu-id="4d06e-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="4d06e-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="4d06e-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4d06e-155">类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-155">Type</span></span>

*   <span data-ttu-id="4d06e-156">String</span><span class="sxs-lookup"><span data-stu-id="4d06e-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4d06e-157">属性</span><span class="sxs-lookup"><span data-stu-id="4d06e-157">Properties</span></span>

|<span data-ttu-id="4d06e-158">名称</span><span class="sxs-lookup"><span data-stu-id="4d06e-158">Name</span></span>| <span data-ttu-id="4d06e-159">类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-159">Type</span></span>| <span data-ttu-id="4d06e-160">描述</span><span class="sxs-lookup"><span data-stu-id="4d06e-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4d06e-161">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-161">String</span></span>|<span data-ttu-id="4d06e-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="4d06e-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4d06e-163">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-163">String</span></span>|<span data-ttu-id="4d06e-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="4d06e-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d06e-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="4d06e-165">Requirements</span></span>

|<span data-ttu-id="4d06e-166">要求</span><span class="sxs-lookup"><span data-stu-id="4d06e-166">Requirement</span></span>| <span data-ttu-id="4d06e-167">值</span><span class="sxs-lookup"><span data-stu-id="4d06e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d06e-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4d06e-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4d06e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="4d06e-169">1.1</span></span>|
|[<span data-ttu-id="4d06e-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4d06e-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4d06e-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="4d06e-172">CoercionType：String</span><span class="sxs-lookup"><span data-stu-id="4d06e-172">CoercionType: String</span></span>

<span data-ttu-id="4d06e-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="4d06e-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4d06e-174">类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-174">Type</span></span>

*   <span data-ttu-id="4d06e-175">String</span><span class="sxs-lookup"><span data-stu-id="4d06e-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4d06e-176">属性</span><span class="sxs-lookup"><span data-stu-id="4d06e-176">Properties</span></span>

|<span data-ttu-id="4d06e-177">名称</span><span class="sxs-lookup"><span data-stu-id="4d06e-177">Name</span></span>| <span data-ttu-id="4d06e-178">类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-178">Type</span></span>| <span data-ttu-id="4d06e-179">描述</span><span class="sxs-lookup"><span data-stu-id="4d06e-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4d06e-180">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-180">String</span></span>|<span data-ttu-id="4d06e-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="4d06e-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4d06e-182">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-182">String</span></span>|<span data-ttu-id="4d06e-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="4d06e-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d06e-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="4d06e-184">Requirements</span></span>

|<span data-ttu-id="4d06e-185">要求</span><span class="sxs-lookup"><span data-stu-id="4d06e-185">Requirement</span></span>| <span data-ttu-id="4d06e-186">值</span><span class="sxs-lookup"><span data-stu-id="4d06e-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d06e-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4d06e-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4d06e-188">1.1</span><span class="sxs-lookup"><span data-stu-id="4d06e-188">1.1</span></span>|
|[<span data-ttu-id="4d06e-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4d06e-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4d06e-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="4d06e-191">EventType：String</span><span class="sxs-lookup"><span data-stu-id="4d06e-191">EventType: String</span></span>

<span data-ttu-id="4d06e-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="4d06e-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4d06e-193">类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-193">Type</span></span>

*   <span data-ttu-id="4d06e-194">String</span><span class="sxs-lookup"><span data-stu-id="4d06e-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4d06e-195">属性</span><span class="sxs-lookup"><span data-stu-id="4d06e-195">Properties</span></span>

| <span data-ttu-id="4d06e-196">名称</span><span class="sxs-lookup"><span data-stu-id="4d06e-196">Name</span></span> | <span data-ttu-id="4d06e-197">类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-197">Type</span></span> | <span data-ttu-id="4d06e-198">描述</span><span class="sxs-lookup"><span data-stu-id="4d06e-198">Description</span></span> | <span data-ttu-id="4d06e-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="4d06e-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="4d06e-200">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-200">String</span></span> | <span data-ttu-id="4d06e-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="4d06e-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4d06e-202">1.7</span><span class="sxs-lookup"><span data-stu-id="4d06e-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="4d06e-203">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-203">String</span></span> | <span data-ttu-id="4d06e-204">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="4d06e-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="4d06e-205">1.8</span><span class="sxs-lookup"><span data-stu-id="4d06e-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="4d06e-206">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-206">String</span></span> | <span data-ttu-id="4d06e-207">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="4d06e-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="4d06e-208">1.8</span><span class="sxs-lookup"><span data-stu-id="4d06e-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="4d06e-209">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-209">String</span></span> | <span data-ttu-id="4d06e-210">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="4d06e-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="4d06e-211">1.5</span><span class="sxs-lookup"><span data-stu-id="4d06e-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4d06e-212">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-212">String</span></span> | <span data-ttu-id="4d06e-213">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="4d06e-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4d06e-214">1.7</span><span class="sxs-lookup"><span data-stu-id="4d06e-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4d06e-215">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-215">String</span></span> | <span data-ttu-id="4d06e-216">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="4d06e-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4d06e-217">1.7</span><span class="sxs-lookup"><span data-stu-id="4d06e-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4d06e-218">Requirements</span><span class="sxs-lookup"><span data-stu-id="4d06e-218">Requirements</span></span>

|<span data-ttu-id="4d06e-219">要求</span><span class="sxs-lookup"><span data-stu-id="4d06e-219">Requirement</span></span>| <span data-ttu-id="4d06e-220">值</span><span class="sxs-lookup"><span data-stu-id="4d06e-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d06e-221">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4d06e-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4d06e-222">1.5</span><span class="sxs-lookup"><span data-stu-id="4d06e-222">1.5</span></span> |
|[<span data-ttu-id="4d06e-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4d06e-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4d06e-224">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="4d06e-225">SourceProperty：String</span><span class="sxs-lookup"><span data-stu-id="4d06e-225">SourceProperty: String</span></span>

<span data-ttu-id="4d06e-226">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="4d06e-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4d06e-227">类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-227">Type</span></span>

*   <span data-ttu-id="4d06e-228">String</span><span class="sxs-lookup"><span data-stu-id="4d06e-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4d06e-229">属性</span><span class="sxs-lookup"><span data-stu-id="4d06e-229">Properties</span></span>

|<span data-ttu-id="4d06e-230">名称</span><span class="sxs-lookup"><span data-stu-id="4d06e-230">Name</span></span>| <span data-ttu-id="4d06e-231">类型</span><span class="sxs-lookup"><span data-stu-id="4d06e-231">Type</span></span>| <span data-ttu-id="4d06e-232">描述</span><span class="sxs-lookup"><span data-stu-id="4d06e-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4d06e-233">字符串</span><span class="sxs-lookup"><span data-stu-id="4d06e-233">String</span></span>|<span data-ttu-id="4d06e-234">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="4d06e-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4d06e-235">String</span><span class="sxs-lookup"><span data-stu-id="4d06e-235">String</span></span>|<span data-ttu-id="4d06e-236">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="4d06e-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4d06e-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="4d06e-237">Requirements</span></span>

|<span data-ttu-id="4d06e-238">要求</span><span class="sxs-lookup"><span data-stu-id="4d06e-238">Requirement</span></span>| <span data-ttu-id="4d06e-239">值</span><span class="sxs-lookup"><span data-stu-id="4d06e-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="4d06e-240">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4d06e-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4d06e-241">1.1</span><span class="sxs-lookup"><span data-stu-id="4d06e-241">1.1</span></span>|
|[<span data-ttu-id="4d06e-242">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4d06e-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4d06e-243">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4d06e-243">Compose or Read</span></span>|
