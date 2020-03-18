---
title: Office 命名空间-要求集1。8
description: Office 命名空间为 Outlook Office 外接程序提供共享接口（要求集1.8）
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0bbe212b0b8e5dc1348cb5cdc03509c44a716d1a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717500"
---
# <a name="office"></a><span data-ttu-id="9d486-103">Office</span><span class="sxs-lookup"><span data-stu-id="9d486-103">Office</span></span>

<span data-ttu-id="9d486-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="9d486-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9d486-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="9d486-106">Requirements</span></span>

|<span data-ttu-id="9d486-107">要求</span><span class="sxs-lookup"><span data-stu-id="9d486-107">Requirement</span></span>| <span data-ttu-id="9d486-108">值</span><span class="sxs-lookup"><span data-stu-id="9d486-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d486-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9d486-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d486-110">1.1</span><span class="sxs-lookup"><span data-stu-id="9d486-110">1.1</span></span>|
|[<span data-ttu-id="9d486-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9d486-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d486-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9d486-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="9d486-113">属性</span><span class="sxs-lookup"><span data-stu-id="9d486-113">Properties</span></span>

| <span data-ttu-id="9d486-114">属性</span><span class="sxs-lookup"><span data-stu-id="9d486-114">Property</span></span> | <span data-ttu-id="9d486-115">型号</span><span class="sxs-lookup"><span data-stu-id="9d486-115">Modes</span></span> | <span data-ttu-id="9d486-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="9d486-116">Return type</span></span> | <span data-ttu-id="9d486-117">最低</span><span class="sxs-lookup"><span data-stu-id="9d486-117">Minimum</span></span><br><span data-ttu-id="9d486-118">要求集</span><span class="sxs-lookup"><span data-stu-id="9d486-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9d486-119">context</span><span class="sxs-lookup"><span data-stu-id="9d486-119">context</span></span>](office.context.md) | <span data-ttu-id="9d486-120">撰写</span><span class="sxs-lookup"><span data-stu-id="9d486-120">Compose</span></span><br><span data-ttu-id="9d486-121">读取</span><span class="sxs-lookup"><span data-stu-id="9d486-121">Read</span></span> | [<span data-ttu-id="9d486-122">Context</span><span class="sxs-lookup"><span data-stu-id="9d486-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="9d486-123">1.1</span><span class="sxs-lookup"><span data-stu-id="9d486-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="9d486-124">枚举</span><span class="sxs-lookup"><span data-stu-id="9d486-124">Enumerations</span></span>

| <span data-ttu-id="9d486-125">枚举</span><span class="sxs-lookup"><span data-stu-id="9d486-125">Enumeration</span></span> | <span data-ttu-id="9d486-126">型号</span><span class="sxs-lookup"><span data-stu-id="9d486-126">Modes</span></span> | <span data-ttu-id="9d486-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="9d486-127">Return type</span></span> | <span data-ttu-id="9d486-128">最低</span><span class="sxs-lookup"><span data-stu-id="9d486-128">Minimum</span></span><br><span data-ttu-id="9d486-129">要求集</span><span class="sxs-lookup"><span data-stu-id="9d486-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9d486-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9d486-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9d486-131">撰写</span><span class="sxs-lookup"><span data-stu-id="9d486-131">Compose</span></span><br><span data-ttu-id="9d486-132">读取</span><span class="sxs-lookup"><span data-stu-id="9d486-132">Read</span></span> | <span data-ttu-id="9d486-133">String</span><span class="sxs-lookup"><span data-stu-id="9d486-133">String</span></span> | [<span data-ttu-id="9d486-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9d486-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9d486-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9d486-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9d486-136">撰写</span><span class="sxs-lookup"><span data-stu-id="9d486-136">Compose</span></span><br><span data-ttu-id="9d486-137">读取</span><span class="sxs-lookup"><span data-stu-id="9d486-137">Read</span></span> | <span data-ttu-id="9d486-138">String</span><span class="sxs-lookup"><span data-stu-id="9d486-138">String</span></span> | [<span data-ttu-id="9d486-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9d486-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9d486-140">EventType</span><span class="sxs-lookup"><span data-stu-id="9d486-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9d486-141">撰写</span><span class="sxs-lookup"><span data-stu-id="9d486-141">Compose</span></span><br><span data-ttu-id="9d486-142">读取</span><span class="sxs-lookup"><span data-stu-id="9d486-142">Read</span></span> | <span data-ttu-id="9d486-143">String</span><span class="sxs-lookup"><span data-stu-id="9d486-143">String</span></span> | [<span data-ttu-id="9d486-144">1.5</span><span class="sxs-lookup"><span data-stu-id="9d486-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9d486-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9d486-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9d486-146">撰写</span><span class="sxs-lookup"><span data-stu-id="9d486-146">Compose</span></span><br><span data-ttu-id="9d486-147">读取</span><span class="sxs-lookup"><span data-stu-id="9d486-147">Read</span></span> | <span data-ttu-id="9d486-148">String</span><span class="sxs-lookup"><span data-stu-id="9d486-148">String</span></span> | [<span data-ttu-id="9d486-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9d486-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="9d486-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="9d486-150">Namespaces</span></span>

<span data-ttu-id="9d486-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="9d486-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="9d486-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="9d486-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="9d486-153">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="9d486-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="9d486-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="9d486-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9d486-155">类型</span><span class="sxs-lookup"><span data-stu-id="9d486-155">Type</span></span>

*   <span data-ttu-id="9d486-156">String</span><span class="sxs-lookup"><span data-stu-id="9d486-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9d486-157">属性：</span><span class="sxs-lookup"><span data-stu-id="9d486-157">Properties:</span></span>

|<span data-ttu-id="9d486-158">姓名</span><span class="sxs-lookup"><span data-stu-id="9d486-158">Name</span></span>| <span data-ttu-id="9d486-159">类型</span><span class="sxs-lookup"><span data-stu-id="9d486-159">Type</span></span>| <span data-ttu-id="9d486-160">说明</span><span class="sxs-lookup"><span data-stu-id="9d486-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9d486-161">String</span><span class="sxs-lookup"><span data-stu-id="9d486-161">String</span></span>|<span data-ttu-id="9d486-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="9d486-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9d486-163">String</span><span class="sxs-lookup"><span data-stu-id="9d486-163">String</span></span>|<span data-ttu-id="9d486-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="9d486-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9d486-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="9d486-165">Requirements</span></span>

|<span data-ttu-id="9d486-166">要求</span><span class="sxs-lookup"><span data-stu-id="9d486-166">Requirement</span></span>| <span data-ttu-id="9d486-167">值</span><span class="sxs-lookup"><span data-stu-id="9d486-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d486-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9d486-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d486-169">1.1</span><span class="sxs-lookup"><span data-stu-id="9d486-169">1.1</span></span>|
|[<span data-ttu-id="9d486-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9d486-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d486-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9d486-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="9d486-172">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="9d486-172">CoercionType: String</span></span>

<span data-ttu-id="9d486-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="9d486-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9d486-174">类型</span><span class="sxs-lookup"><span data-stu-id="9d486-174">Type</span></span>

*   <span data-ttu-id="9d486-175">String</span><span class="sxs-lookup"><span data-stu-id="9d486-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9d486-176">属性：</span><span class="sxs-lookup"><span data-stu-id="9d486-176">Properties:</span></span>

|<span data-ttu-id="9d486-177">姓名</span><span class="sxs-lookup"><span data-stu-id="9d486-177">Name</span></span>| <span data-ttu-id="9d486-178">类型</span><span class="sxs-lookup"><span data-stu-id="9d486-178">Type</span></span>| <span data-ttu-id="9d486-179">说明</span><span class="sxs-lookup"><span data-stu-id="9d486-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9d486-180">String</span><span class="sxs-lookup"><span data-stu-id="9d486-180">String</span></span>|<span data-ttu-id="9d486-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="9d486-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9d486-182">String</span><span class="sxs-lookup"><span data-stu-id="9d486-182">String</span></span>|<span data-ttu-id="9d486-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="9d486-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9d486-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="9d486-184">Requirements</span></span>

|<span data-ttu-id="9d486-185">要求</span><span class="sxs-lookup"><span data-stu-id="9d486-185">Requirement</span></span>| <span data-ttu-id="9d486-186">值</span><span class="sxs-lookup"><span data-stu-id="9d486-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d486-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9d486-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d486-188">1.1</span><span class="sxs-lookup"><span data-stu-id="9d486-188">1.1</span></span>|
|[<span data-ttu-id="9d486-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9d486-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d486-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9d486-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="9d486-191">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="9d486-191">EventType: String</span></span>

<span data-ttu-id="9d486-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="9d486-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9d486-193">类型</span><span class="sxs-lookup"><span data-stu-id="9d486-193">Type</span></span>

*   <span data-ttu-id="9d486-194">String</span><span class="sxs-lookup"><span data-stu-id="9d486-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9d486-195">属性：</span><span class="sxs-lookup"><span data-stu-id="9d486-195">Properties:</span></span>

| <span data-ttu-id="9d486-196">姓名</span><span class="sxs-lookup"><span data-stu-id="9d486-196">Name</span></span> | <span data-ttu-id="9d486-197">类型</span><span class="sxs-lookup"><span data-stu-id="9d486-197">Type</span></span> | <span data-ttu-id="9d486-198">说明</span><span class="sxs-lookup"><span data-stu-id="9d486-198">Description</span></span> | <span data-ttu-id="9d486-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="9d486-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="9d486-200">String</span><span class="sxs-lookup"><span data-stu-id="9d486-200">String</span></span> | <span data-ttu-id="9d486-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="9d486-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="9d486-202">1.7</span><span class="sxs-lookup"><span data-stu-id="9d486-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="9d486-203">String</span><span class="sxs-lookup"><span data-stu-id="9d486-203">String</span></span> | <span data-ttu-id="9d486-204">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="9d486-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="9d486-205">1.8</span><span class="sxs-lookup"><span data-stu-id="9d486-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="9d486-206">String</span><span class="sxs-lookup"><span data-stu-id="9d486-206">String</span></span> | <span data-ttu-id="9d486-207">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="9d486-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="9d486-208">1.8</span><span class="sxs-lookup"><span data-stu-id="9d486-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="9d486-209">String</span><span class="sxs-lookup"><span data-stu-id="9d486-209">String</span></span> | <span data-ttu-id="9d486-210">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="9d486-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="9d486-211">1.5</span><span class="sxs-lookup"><span data-stu-id="9d486-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="9d486-212">String</span><span class="sxs-lookup"><span data-stu-id="9d486-212">String</span></span> | <span data-ttu-id="9d486-213">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="9d486-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="9d486-214">1.7</span><span class="sxs-lookup"><span data-stu-id="9d486-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="9d486-215">String</span><span class="sxs-lookup"><span data-stu-id="9d486-215">String</span></span> | <span data-ttu-id="9d486-216">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="9d486-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="9d486-217">1.7</span><span class="sxs-lookup"><span data-stu-id="9d486-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9d486-218">Requirements</span><span class="sxs-lookup"><span data-stu-id="9d486-218">Requirements</span></span>

|<span data-ttu-id="9d486-219">要求</span><span class="sxs-lookup"><span data-stu-id="9d486-219">Requirement</span></span>| <span data-ttu-id="9d486-220">值</span><span class="sxs-lookup"><span data-stu-id="9d486-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d486-221">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9d486-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d486-222">1.5</span><span class="sxs-lookup"><span data-stu-id="9d486-222">1.5</span></span> |
|[<span data-ttu-id="9d486-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9d486-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d486-224">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9d486-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="9d486-225">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="9d486-225">SourceProperty: String</span></span>

<span data-ttu-id="9d486-226">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="9d486-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9d486-227">类型</span><span class="sxs-lookup"><span data-stu-id="9d486-227">Type</span></span>

*   <span data-ttu-id="9d486-228">String</span><span class="sxs-lookup"><span data-stu-id="9d486-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9d486-229">属性：</span><span class="sxs-lookup"><span data-stu-id="9d486-229">Properties:</span></span>

|<span data-ttu-id="9d486-230">姓名</span><span class="sxs-lookup"><span data-stu-id="9d486-230">Name</span></span>| <span data-ttu-id="9d486-231">类型</span><span class="sxs-lookup"><span data-stu-id="9d486-231">Type</span></span>| <span data-ttu-id="9d486-232">说明</span><span class="sxs-lookup"><span data-stu-id="9d486-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9d486-233">String</span><span class="sxs-lookup"><span data-stu-id="9d486-233">String</span></span>|<span data-ttu-id="9d486-234">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="9d486-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9d486-235">String</span><span class="sxs-lookup"><span data-stu-id="9d486-235">String</span></span>|<span data-ttu-id="9d486-236">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="9d486-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9d486-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="9d486-237">Requirements</span></span>

|<span data-ttu-id="9d486-238">要求</span><span class="sxs-lookup"><span data-stu-id="9d486-238">Requirement</span></span>| <span data-ttu-id="9d486-239">值</span><span class="sxs-lookup"><span data-stu-id="9d486-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d486-240">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9d486-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d486-241">1.1</span><span class="sxs-lookup"><span data-stu-id="9d486-241">1.1</span></span>|
|[<span data-ttu-id="9d486-242">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9d486-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d486-243">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9d486-243">Compose or Read</span></span>|
