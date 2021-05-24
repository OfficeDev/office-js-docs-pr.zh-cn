---
title: Office命名空间 - 要求集 1.7
description: Office邮箱 API 要求集 1.7 Outlook外接程序可用的命名空间成员。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 19c80c0c8c4aaf31c42aad16b3f474e92b7cdaec
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590972"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="d96be-103">Office (邮箱要求集 1.7) </span><span class="sxs-lookup"><span data-stu-id="d96be-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="d96be-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="d96be-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d96be-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="d96be-106">Requirements</span></span>

|<span data-ttu-id="d96be-107">要求</span><span class="sxs-lookup"><span data-stu-id="d96be-107">Requirement</span></span>| <span data-ttu-id="d96be-108">值</span><span class="sxs-lookup"><span data-stu-id="d96be-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d96be-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d96be-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d96be-110">1.1</span><span class="sxs-lookup"><span data-stu-id="d96be-110">1.1</span></span>|
|[<span data-ttu-id="d96be-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d96be-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d96be-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="d96be-113">属性</span><span class="sxs-lookup"><span data-stu-id="d96be-113">Properties</span></span>

| <span data-ttu-id="d96be-114">属性</span><span class="sxs-lookup"><span data-stu-id="d96be-114">Property</span></span> | <span data-ttu-id="d96be-115">模式</span><span class="sxs-lookup"><span data-stu-id="d96be-115">Modes</span></span> | <span data-ttu-id="d96be-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="d96be-116">Return type</span></span> | <span data-ttu-id="d96be-117">最小值</span><span class="sxs-lookup"><span data-stu-id="d96be-117">Minimum</span></span><br><span data-ttu-id="d96be-118">要求集</span><span class="sxs-lookup"><span data-stu-id="d96be-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d96be-119">context</span><span class="sxs-lookup"><span data-stu-id="d96be-119">context</span></span>](office.context.md) | <span data-ttu-id="d96be-120">撰写</span><span class="sxs-lookup"><span data-stu-id="d96be-120">Compose</span></span><br><span data-ttu-id="d96be-121">阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-121">Read</span></span> | [<span data-ttu-id="d96be-122">Context</span><span class="sxs-lookup"><span data-stu-id="d96be-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="d96be-123">1.1</span><span class="sxs-lookup"><span data-stu-id="d96be-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="d96be-124">枚举</span><span class="sxs-lookup"><span data-stu-id="d96be-124">Enumerations</span></span>

| <span data-ttu-id="d96be-125">枚举</span><span class="sxs-lookup"><span data-stu-id="d96be-125">Enumeration</span></span> | <span data-ttu-id="d96be-126">模式</span><span class="sxs-lookup"><span data-stu-id="d96be-126">Modes</span></span> | <span data-ttu-id="d96be-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="d96be-127">Return type</span></span> | <span data-ttu-id="d96be-128">最小值</span><span class="sxs-lookup"><span data-stu-id="d96be-128">Minimum</span></span><br><span data-ttu-id="d96be-129">要求集</span><span class="sxs-lookup"><span data-stu-id="d96be-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d96be-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d96be-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d96be-131">撰写</span><span class="sxs-lookup"><span data-stu-id="d96be-131">Compose</span></span><br><span data-ttu-id="d96be-132">阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-132">Read</span></span> | <span data-ttu-id="d96be-133">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-133">String</span></span> | [<span data-ttu-id="d96be-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d96be-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d96be-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d96be-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d96be-136">撰写</span><span class="sxs-lookup"><span data-stu-id="d96be-136">Compose</span></span><br><span data-ttu-id="d96be-137">阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-137">Read</span></span> | <span data-ttu-id="d96be-138">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-138">String</span></span> | [<span data-ttu-id="d96be-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d96be-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d96be-140">EventType</span><span class="sxs-lookup"><span data-stu-id="d96be-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d96be-141">撰写</span><span class="sxs-lookup"><span data-stu-id="d96be-141">Compose</span></span><br><span data-ttu-id="d96be-142">阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-142">Read</span></span> | <span data-ttu-id="d96be-143">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-143">String</span></span> | [<span data-ttu-id="d96be-144">1.5</span><span class="sxs-lookup"><span data-stu-id="d96be-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="d96be-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d96be-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d96be-146">撰写</span><span class="sxs-lookup"><span data-stu-id="d96be-146">Compose</span></span><br><span data-ttu-id="d96be-147">阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-147">Read</span></span> | <span data-ttu-id="d96be-148">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-148">String</span></span> | [<span data-ttu-id="d96be-149">1.1</span><span class="sxs-lookup"><span data-stu-id="d96be-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="d96be-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="d96be-150">Namespaces</span></span>

<span data-ttu-id="d96be-151">[MailboxEnums：](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true)包括许多Outlook枚举，例如、 `ItemType` `EntityType` 和 `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="d96be-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="d96be-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="d96be-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d96be-153">AsyncResultStatus：String</span><span class="sxs-lookup"><span data-stu-id="d96be-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="d96be-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="d96be-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d96be-155">类型</span><span class="sxs-lookup"><span data-stu-id="d96be-155">Type</span></span>

*   <span data-ttu-id="d96be-156">String</span><span class="sxs-lookup"><span data-stu-id="d96be-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d96be-157">属性</span><span class="sxs-lookup"><span data-stu-id="d96be-157">Properties</span></span>

|<span data-ttu-id="d96be-158">名称</span><span class="sxs-lookup"><span data-stu-id="d96be-158">Name</span></span>| <span data-ttu-id="d96be-159">类型</span><span class="sxs-lookup"><span data-stu-id="d96be-159">Type</span></span>| <span data-ttu-id="d96be-160">描述</span><span class="sxs-lookup"><span data-stu-id="d96be-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d96be-161">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-161">String</span></span>|<span data-ttu-id="d96be-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="d96be-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d96be-163">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-163">String</span></span>|<span data-ttu-id="d96be-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="d96be-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d96be-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="d96be-165">Requirements</span></span>

|<span data-ttu-id="d96be-166">要求</span><span class="sxs-lookup"><span data-stu-id="d96be-166">Requirement</span></span>| <span data-ttu-id="d96be-167">值</span><span class="sxs-lookup"><span data-stu-id="d96be-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d96be-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d96be-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d96be-169">1.1</span><span class="sxs-lookup"><span data-stu-id="d96be-169">1.1</span></span>|
|[<span data-ttu-id="d96be-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d96be-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d96be-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d96be-172">CoercionType：String</span><span class="sxs-lookup"><span data-stu-id="d96be-172">CoercionType: String</span></span>

<span data-ttu-id="d96be-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="d96be-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d96be-174">类型</span><span class="sxs-lookup"><span data-stu-id="d96be-174">Type</span></span>

*   <span data-ttu-id="d96be-175">String</span><span class="sxs-lookup"><span data-stu-id="d96be-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d96be-176">属性</span><span class="sxs-lookup"><span data-stu-id="d96be-176">Properties</span></span>

|<span data-ttu-id="d96be-177">名称</span><span class="sxs-lookup"><span data-stu-id="d96be-177">Name</span></span>| <span data-ttu-id="d96be-178">类型</span><span class="sxs-lookup"><span data-stu-id="d96be-178">Type</span></span>| <span data-ttu-id="d96be-179">描述</span><span class="sxs-lookup"><span data-stu-id="d96be-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d96be-180">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-180">String</span></span>|<span data-ttu-id="d96be-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d96be-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d96be-182">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-182">String</span></span>|<span data-ttu-id="d96be-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d96be-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d96be-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="d96be-184">Requirements</span></span>

|<span data-ttu-id="d96be-185">要求</span><span class="sxs-lookup"><span data-stu-id="d96be-185">Requirement</span></span>| <span data-ttu-id="d96be-186">值</span><span class="sxs-lookup"><span data-stu-id="d96be-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="d96be-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d96be-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d96be-188">1.1</span><span class="sxs-lookup"><span data-stu-id="d96be-188">1.1</span></span>|
|[<span data-ttu-id="d96be-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d96be-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d96be-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="d96be-191">EventType：String</span><span class="sxs-lookup"><span data-stu-id="d96be-191">EventType: String</span></span>

<span data-ttu-id="d96be-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="d96be-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d96be-193">类型</span><span class="sxs-lookup"><span data-stu-id="d96be-193">Type</span></span>

*   <span data-ttu-id="d96be-194">String</span><span class="sxs-lookup"><span data-stu-id="d96be-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d96be-195">属性</span><span class="sxs-lookup"><span data-stu-id="d96be-195">Properties</span></span>

| <span data-ttu-id="d96be-196">名称</span><span class="sxs-lookup"><span data-stu-id="d96be-196">Name</span></span> | <span data-ttu-id="d96be-197">类型</span><span class="sxs-lookup"><span data-stu-id="d96be-197">Type</span></span> | <span data-ttu-id="d96be-198">描述</span><span class="sxs-lookup"><span data-stu-id="d96be-198">Description</span></span> | <span data-ttu-id="d96be-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="d96be-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="d96be-200">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-200">String</span></span> | <span data-ttu-id="d96be-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="d96be-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="d96be-202">1.7</span><span class="sxs-lookup"><span data-stu-id="d96be-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="d96be-203">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-203">String</span></span> | <span data-ttu-id="d96be-204">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="d96be-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="d96be-205">1.5</span><span class="sxs-lookup"><span data-stu-id="d96be-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="d96be-206">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-206">String</span></span> | <span data-ttu-id="d96be-207">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="d96be-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="d96be-208">1.7</span><span class="sxs-lookup"><span data-stu-id="d96be-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="d96be-209">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-209">String</span></span> | <span data-ttu-id="d96be-210">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="d96be-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="d96be-211">1.7</span><span class="sxs-lookup"><span data-stu-id="d96be-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d96be-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="d96be-212">Requirements</span></span>

|<span data-ttu-id="d96be-213">要求</span><span class="sxs-lookup"><span data-stu-id="d96be-213">Requirement</span></span>| <span data-ttu-id="d96be-214">值</span><span class="sxs-lookup"><span data-stu-id="d96be-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="d96be-215">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d96be-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d96be-216">1.5</span><span class="sxs-lookup"><span data-stu-id="d96be-216">1.5</span></span> |
|[<span data-ttu-id="d96be-217">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d96be-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d96be-218">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d96be-219">SourceProperty：String</span><span class="sxs-lookup"><span data-stu-id="d96be-219">SourceProperty: String</span></span>

<span data-ttu-id="d96be-220">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="d96be-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d96be-221">类型</span><span class="sxs-lookup"><span data-stu-id="d96be-221">Type</span></span>

*   <span data-ttu-id="d96be-222">String</span><span class="sxs-lookup"><span data-stu-id="d96be-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d96be-223">属性</span><span class="sxs-lookup"><span data-stu-id="d96be-223">Properties</span></span>

|<span data-ttu-id="d96be-224">名称</span><span class="sxs-lookup"><span data-stu-id="d96be-224">Name</span></span>| <span data-ttu-id="d96be-225">类型</span><span class="sxs-lookup"><span data-stu-id="d96be-225">Type</span></span>| <span data-ttu-id="d96be-226">描述</span><span class="sxs-lookup"><span data-stu-id="d96be-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d96be-227">字符串</span><span class="sxs-lookup"><span data-stu-id="d96be-227">String</span></span>|<span data-ttu-id="d96be-228">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="d96be-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d96be-229">String</span><span class="sxs-lookup"><span data-stu-id="d96be-229">String</span></span>|<span data-ttu-id="d96be-230">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="d96be-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d96be-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="d96be-231">Requirements</span></span>

|<span data-ttu-id="d96be-232">要求</span><span class="sxs-lookup"><span data-stu-id="d96be-232">Requirement</span></span>| <span data-ttu-id="d96be-233">值</span><span class="sxs-lookup"><span data-stu-id="d96be-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="d96be-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d96be-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d96be-235">1.1</span><span class="sxs-lookup"><span data-stu-id="d96be-235">1.1</span></span>|
|[<span data-ttu-id="d96be-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d96be-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d96be-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d96be-237">Compose or Read</span></span>|
