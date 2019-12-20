---
title: Office 命名空间-要求集1。8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: b23afd7b84dcd18e120f6aea4bd4fb0952791f1c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814164"
---
# <a name="office"></a><span data-ttu-id="69348-102">Office</span><span class="sxs-lookup"><span data-stu-id="69348-102">Office</span></span>

<span data-ttu-id="69348-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="69348-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="69348-105">要求</span><span class="sxs-lookup"><span data-stu-id="69348-105">Requirements</span></span>

|<span data-ttu-id="69348-106">要求</span><span class="sxs-lookup"><span data-stu-id="69348-106">Requirement</span></span>| <span data-ttu-id="69348-107">值</span><span class="sxs-lookup"><span data-stu-id="69348-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="69348-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="69348-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="69348-109">1.1</span><span class="sxs-lookup"><span data-stu-id="69348-109">1.1</span></span>|
|[<span data-ttu-id="69348-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="69348-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="69348-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="69348-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="69348-112">属性</span><span class="sxs-lookup"><span data-stu-id="69348-112">Properties</span></span>

| <span data-ttu-id="69348-113">属性</span><span class="sxs-lookup"><span data-stu-id="69348-113">Property</span></span> | <span data-ttu-id="69348-114">型号</span><span class="sxs-lookup"><span data-stu-id="69348-114">Modes</span></span> | <span data-ttu-id="69348-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="69348-115">Return type</span></span> | <span data-ttu-id="69348-116">最低</span><span class="sxs-lookup"><span data-stu-id="69348-116">Minimum</span></span><br><span data-ttu-id="69348-117">要求集</span><span class="sxs-lookup"><span data-stu-id="69348-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="69348-118">context</span><span class="sxs-lookup"><span data-stu-id="69348-118">context</span></span>](office.context.md) | <span data-ttu-id="69348-119">撰写</span><span class="sxs-lookup"><span data-stu-id="69348-119">Compose</span></span><br><span data-ttu-id="69348-120">读取</span><span class="sxs-lookup"><span data-stu-id="69348-120">Read</span></span> | [<span data-ttu-id="69348-121">Context</span><span class="sxs-lookup"><span data-stu-id="69348-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="69348-122">1.1</span><span class="sxs-lookup"><span data-stu-id="69348-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="69348-123">枚举</span><span class="sxs-lookup"><span data-stu-id="69348-123">Enumerations</span></span>

| <span data-ttu-id="69348-124">枚举</span><span class="sxs-lookup"><span data-stu-id="69348-124">Enumeration</span></span> | <span data-ttu-id="69348-125">型号</span><span class="sxs-lookup"><span data-stu-id="69348-125">Modes</span></span> | <span data-ttu-id="69348-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="69348-126">Return type</span></span> | <span data-ttu-id="69348-127">最低</span><span class="sxs-lookup"><span data-stu-id="69348-127">Minimum</span></span><br><span data-ttu-id="69348-128">要求集</span><span class="sxs-lookup"><span data-stu-id="69348-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="69348-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="69348-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="69348-130">撰写</span><span class="sxs-lookup"><span data-stu-id="69348-130">Compose</span></span><br><span data-ttu-id="69348-131">读取</span><span class="sxs-lookup"><span data-stu-id="69348-131">Read</span></span> | <span data-ttu-id="69348-132">String</span><span class="sxs-lookup"><span data-stu-id="69348-132">String</span></span> | [<span data-ttu-id="69348-133">1.1</span><span class="sxs-lookup"><span data-stu-id="69348-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="69348-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="69348-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="69348-135">撰写</span><span class="sxs-lookup"><span data-stu-id="69348-135">Compose</span></span><br><span data-ttu-id="69348-136">读取</span><span class="sxs-lookup"><span data-stu-id="69348-136">Read</span></span> | <span data-ttu-id="69348-137">String</span><span class="sxs-lookup"><span data-stu-id="69348-137">String</span></span> | [<span data-ttu-id="69348-138">1.1</span><span class="sxs-lookup"><span data-stu-id="69348-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="69348-139">EventType</span><span class="sxs-lookup"><span data-stu-id="69348-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="69348-140">撰写</span><span class="sxs-lookup"><span data-stu-id="69348-140">Compose</span></span><br><span data-ttu-id="69348-141">读取</span><span class="sxs-lookup"><span data-stu-id="69348-141">Read</span></span> | <span data-ttu-id="69348-142">String</span><span class="sxs-lookup"><span data-stu-id="69348-142">String</span></span> | [<span data-ttu-id="69348-143">1.5</span><span class="sxs-lookup"><span data-stu-id="69348-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="69348-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="69348-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="69348-145">撰写</span><span class="sxs-lookup"><span data-stu-id="69348-145">Compose</span></span><br><span data-ttu-id="69348-146">读取</span><span class="sxs-lookup"><span data-stu-id="69348-146">Read</span></span> | <span data-ttu-id="69348-147">String</span><span class="sxs-lookup"><span data-stu-id="69348-147">String</span></span> | [<span data-ttu-id="69348-148">1.1</span><span class="sxs-lookup"><span data-stu-id="69348-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="69348-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="69348-149">Namespaces</span></span>

<span data-ttu-id="69348-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="69348-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="69348-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="69348-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="69348-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="69348-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="69348-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="69348-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="69348-154">类型</span><span class="sxs-lookup"><span data-stu-id="69348-154">Type</span></span>

*   <span data-ttu-id="69348-155">String</span><span class="sxs-lookup"><span data-stu-id="69348-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="69348-156">属性：</span><span class="sxs-lookup"><span data-stu-id="69348-156">Properties:</span></span>

|<span data-ttu-id="69348-157">名称</span><span class="sxs-lookup"><span data-stu-id="69348-157">Name</span></span>| <span data-ttu-id="69348-158">类型</span><span class="sxs-lookup"><span data-stu-id="69348-158">Type</span></span>| <span data-ttu-id="69348-159">说明</span><span class="sxs-lookup"><span data-stu-id="69348-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="69348-160">String</span><span class="sxs-lookup"><span data-stu-id="69348-160">String</span></span>|<span data-ttu-id="69348-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="69348-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="69348-162">String</span><span class="sxs-lookup"><span data-stu-id="69348-162">String</span></span>|<span data-ttu-id="69348-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="69348-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="69348-164">要求</span><span class="sxs-lookup"><span data-stu-id="69348-164">Requirements</span></span>

|<span data-ttu-id="69348-165">要求</span><span class="sxs-lookup"><span data-stu-id="69348-165">Requirement</span></span>| <span data-ttu-id="69348-166">值</span><span class="sxs-lookup"><span data-stu-id="69348-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="69348-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="69348-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="69348-168">1.1</span><span class="sxs-lookup"><span data-stu-id="69348-168">1.1</span></span>|
|[<span data-ttu-id="69348-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="69348-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="69348-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="69348-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="69348-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="69348-171">CoercionType: String</span></span>

<span data-ttu-id="69348-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="69348-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="69348-173">类型</span><span class="sxs-lookup"><span data-stu-id="69348-173">Type</span></span>

*   <span data-ttu-id="69348-174">String</span><span class="sxs-lookup"><span data-stu-id="69348-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="69348-175">属性：</span><span class="sxs-lookup"><span data-stu-id="69348-175">Properties:</span></span>

|<span data-ttu-id="69348-176">名称</span><span class="sxs-lookup"><span data-stu-id="69348-176">Name</span></span>| <span data-ttu-id="69348-177">类型</span><span class="sxs-lookup"><span data-stu-id="69348-177">Type</span></span>| <span data-ttu-id="69348-178">说明</span><span class="sxs-lookup"><span data-stu-id="69348-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="69348-179">String</span><span class="sxs-lookup"><span data-stu-id="69348-179">String</span></span>|<span data-ttu-id="69348-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="69348-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="69348-181">String</span><span class="sxs-lookup"><span data-stu-id="69348-181">String</span></span>|<span data-ttu-id="69348-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="69348-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="69348-183">要求</span><span class="sxs-lookup"><span data-stu-id="69348-183">Requirements</span></span>

|<span data-ttu-id="69348-184">要求</span><span class="sxs-lookup"><span data-stu-id="69348-184">Requirement</span></span>| <span data-ttu-id="69348-185">值</span><span class="sxs-lookup"><span data-stu-id="69348-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="69348-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="69348-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="69348-187">1.1</span><span class="sxs-lookup"><span data-stu-id="69348-187">1.1</span></span>|
|[<span data-ttu-id="69348-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="69348-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="69348-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="69348-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="69348-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="69348-190">EventType: String</span></span>

<span data-ttu-id="69348-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="69348-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="69348-192">类型</span><span class="sxs-lookup"><span data-stu-id="69348-192">Type</span></span>

*   <span data-ttu-id="69348-193">String</span><span class="sxs-lookup"><span data-stu-id="69348-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="69348-194">属性：</span><span class="sxs-lookup"><span data-stu-id="69348-194">Properties:</span></span>

| <span data-ttu-id="69348-195">名称</span><span class="sxs-lookup"><span data-stu-id="69348-195">Name</span></span> | <span data-ttu-id="69348-196">类型</span><span class="sxs-lookup"><span data-stu-id="69348-196">Type</span></span> | <span data-ttu-id="69348-197">说明</span><span class="sxs-lookup"><span data-stu-id="69348-197">Description</span></span> | <span data-ttu-id="69348-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="69348-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="69348-199">String</span><span class="sxs-lookup"><span data-stu-id="69348-199">String</span></span> | <span data-ttu-id="69348-200">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="69348-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="69348-201">1.7</span><span class="sxs-lookup"><span data-stu-id="69348-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="69348-202">String</span><span class="sxs-lookup"><span data-stu-id="69348-202">String</span></span> | <span data-ttu-id="69348-203">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="69348-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="69348-204">1.8</span><span class="sxs-lookup"><span data-stu-id="69348-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="69348-205">String</span><span class="sxs-lookup"><span data-stu-id="69348-205">String</span></span> | <span data-ttu-id="69348-206">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="69348-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="69348-207">1.8</span><span class="sxs-lookup"><span data-stu-id="69348-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="69348-208">String</span><span class="sxs-lookup"><span data-stu-id="69348-208">String</span></span> | <span data-ttu-id="69348-209">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="69348-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="69348-210">1.5</span><span class="sxs-lookup"><span data-stu-id="69348-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="69348-211">String</span><span class="sxs-lookup"><span data-stu-id="69348-211">String</span></span> | <span data-ttu-id="69348-212">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="69348-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="69348-213">1.7</span><span class="sxs-lookup"><span data-stu-id="69348-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="69348-214">String</span><span class="sxs-lookup"><span data-stu-id="69348-214">String</span></span> | <span data-ttu-id="69348-215">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="69348-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="69348-216">1.7</span><span class="sxs-lookup"><span data-stu-id="69348-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="69348-217">要求</span><span class="sxs-lookup"><span data-stu-id="69348-217">Requirements</span></span>

|<span data-ttu-id="69348-218">要求</span><span class="sxs-lookup"><span data-stu-id="69348-218">Requirement</span></span>| <span data-ttu-id="69348-219">值</span><span class="sxs-lookup"><span data-stu-id="69348-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="69348-220">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="69348-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="69348-221">1.5</span><span class="sxs-lookup"><span data-stu-id="69348-221">1.5</span></span> |
|[<span data-ttu-id="69348-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="69348-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="69348-223">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="69348-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="69348-224">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="69348-224">SourceProperty: String</span></span>

<span data-ttu-id="69348-225">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="69348-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="69348-226">类型</span><span class="sxs-lookup"><span data-stu-id="69348-226">Type</span></span>

*   <span data-ttu-id="69348-227">String</span><span class="sxs-lookup"><span data-stu-id="69348-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="69348-228">属性：</span><span class="sxs-lookup"><span data-stu-id="69348-228">Properties:</span></span>

|<span data-ttu-id="69348-229">名称</span><span class="sxs-lookup"><span data-stu-id="69348-229">Name</span></span>| <span data-ttu-id="69348-230">类型</span><span class="sxs-lookup"><span data-stu-id="69348-230">Type</span></span>| <span data-ttu-id="69348-231">说明</span><span class="sxs-lookup"><span data-stu-id="69348-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="69348-232">String</span><span class="sxs-lookup"><span data-stu-id="69348-232">String</span></span>|<span data-ttu-id="69348-233">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="69348-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="69348-234">String</span><span class="sxs-lookup"><span data-stu-id="69348-234">String</span></span>|<span data-ttu-id="69348-235">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="69348-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="69348-236">要求</span><span class="sxs-lookup"><span data-stu-id="69348-236">Requirements</span></span>

|<span data-ttu-id="69348-237">要求</span><span class="sxs-lookup"><span data-stu-id="69348-237">Requirement</span></span>| <span data-ttu-id="69348-238">值</span><span class="sxs-lookup"><span data-stu-id="69348-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="69348-239">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="69348-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="69348-240">1.1</span><span class="sxs-lookup"><span data-stu-id="69348-240">1.1</span></span>|
|[<span data-ttu-id="69348-241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="69348-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="69348-242">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="69348-242">Compose or Read</span></span>|
