---
title: Office 命名空间-要求集1。7
description: 此命名空间提供 Outlook Office 外接程序使用的共享接口（要求集1.7）
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 50fa22ac14aee3b7276be83813db248681435dc1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717593"
---
# <a name="office"></a><span data-ttu-id="e021e-103">Office</span><span class="sxs-lookup"><span data-stu-id="e021e-103">Office</span></span>

<span data-ttu-id="e021e-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="e021e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e021e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="e021e-106">Requirements</span></span>

|<span data-ttu-id="e021e-107">要求</span><span class="sxs-lookup"><span data-stu-id="e021e-107">Requirement</span></span>| <span data-ttu-id="e021e-108">值</span><span class="sxs-lookup"><span data-stu-id="e021e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e021e-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e021e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e021e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="e021e-110">1.1</span></span>|
|[<span data-ttu-id="e021e-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e021e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e021e-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e021e-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e021e-113">属性</span><span class="sxs-lookup"><span data-stu-id="e021e-113">Properties</span></span>

| <span data-ttu-id="e021e-114">属性</span><span class="sxs-lookup"><span data-stu-id="e021e-114">Property</span></span> | <span data-ttu-id="e021e-115">型号</span><span class="sxs-lookup"><span data-stu-id="e021e-115">Modes</span></span> | <span data-ttu-id="e021e-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="e021e-116">Return type</span></span> | <span data-ttu-id="e021e-117">最低</span><span class="sxs-lookup"><span data-stu-id="e021e-117">Minimum</span></span><br><span data-ttu-id="e021e-118">要求集</span><span class="sxs-lookup"><span data-stu-id="e021e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e021e-119">context</span><span class="sxs-lookup"><span data-stu-id="e021e-119">context</span></span>](office.context.md) | <span data-ttu-id="e021e-120">撰写</span><span class="sxs-lookup"><span data-stu-id="e021e-120">Compose</span></span><br><span data-ttu-id="e021e-121">读取</span><span class="sxs-lookup"><span data-stu-id="e021e-121">Read</span></span> | [<span data-ttu-id="e021e-122">Context</span><span class="sxs-lookup"><span data-stu-id="e021e-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="e021e-123">1.1</span><span class="sxs-lookup"><span data-stu-id="e021e-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="e021e-124">枚举</span><span class="sxs-lookup"><span data-stu-id="e021e-124">Enumerations</span></span>

| <span data-ttu-id="e021e-125">枚举</span><span class="sxs-lookup"><span data-stu-id="e021e-125">Enumeration</span></span> | <span data-ttu-id="e021e-126">型号</span><span class="sxs-lookup"><span data-stu-id="e021e-126">Modes</span></span> | <span data-ttu-id="e021e-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="e021e-127">Return type</span></span> | <span data-ttu-id="e021e-128">最低</span><span class="sxs-lookup"><span data-stu-id="e021e-128">Minimum</span></span><br><span data-ttu-id="e021e-129">要求集</span><span class="sxs-lookup"><span data-stu-id="e021e-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e021e-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e021e-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e021e-131">撰写</span><span class="sxs-lookup"><span data-stu-id="e021e-131">Compose</span></span><br><span data-ttu-id="e021e-132">读取</span><span class="sxs-lookup"><span data-stu-id="e021e-132">Read</span></span> | <span data-ttu-id="e021e-133">String</span><span class="sxs-lookup"><span data-stu-id="e021e-133">String</span></span> | [<span data-ttu-id="e021e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e021e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e021e-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e021e-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e021e-136">撰写</span><span class="sxs-lookup"><span data-stu-id="e021e-136">Compose</span></span><br><span data-ttu-id="e021e-137">读取</span><span class="sxs-lookup"><span data-stu-id="e021e-137">Read</span></span> | <span data-ttu-id="e021e-138">String</span><span class="sxs-lookup"><span data-stu-id="e021e-138">String</span></span> | [<span data-ttu-id="e021e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e021e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e021e-140">EventType</span><span class="sxs-lookup"><span data-stu-id="e021e-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e021e-141">撰写</span><span class="sxs-lookup"><span data-stu-id="e021e-141">Compose</span></span><br><span data-ttu-id="e021e-142">读取</span><span class="sxs-lookup"><span data-stu-id="e021e-142">Read</span></span> | <span data-ttu-id="e021e-143">String</span><span class="sxs-lookup"><span data-stu-id="e021e-143">String</span></span> | [<span data-ttu-id="e021e-144">1.5</span><span class="sxs-lookup"><span data-stu-id="e021e-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e021e-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e021e-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e021e-146">撰写</span><span class="sxs-lookup"><span data-stu-id="e021e-146">Compose</span></span><br><span data-ttu-id="e021e-147">读取</span><span class="sxs-lookup"><span data-stu-id="e021e-147">Read</span></span> | <span data-ttu-id="e021e-148">String</span><span class="sxs-lookup"><span data-stu-id="e021e-148">String</span></span> | [<span data-ttu-id="e021e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e021e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="e021e-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="e021e-150">Namespaces</span></span>

<span data-ttu-id="e021e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="e021e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="e021e-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="e021e-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e021e-153">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="e021e-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="e021e-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="e021e-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e021e-155">类型</span><span class="sxs-lookup"><span data-stu-id="e021e-155">Type</span></span>

*   <span data-ttu-id="e021e-156">String</span><span class="sxs-lookup"><span data-stu-id="e021e-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e021e-157">属性：</span><span class="sxs-lookup"><span data-stu-id="e021e-157">Properties:</span></span>

|<span data-ttu-id="e021e-158">姓名</span><span class="sxs-lookup"><span data-stu-id="e021e-158">Name</span></span>| <span data-ttu-id="e021e-159">类型</span><span class="sxs-lookup"><span data-stu-id="e021e-159">Type</span></span>| <span data-ttu-id="e021e-160">说明</span><span class="sxs-lookup"><span data-stu-id="e021e-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e021e-161">String</span><span class="sxs-lookup"><span data-stu-id="e021e-161">String</span></span>|<span data-ttu-id="e021e-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="e021e-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e021e-163">String</span><span class="sxs-lookup"><span data-stu-id="e021e-163">String</span></span>|<span data-ttu-id="e021e-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="e021e-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e021e-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="e021e-165">Requirements</span></span>

|<span data-ttu-id="e021e-166">要求</span><span class="sxs-lookup"><span data-stu-id="e021e-166">Requirement</span></span>| <span data-ttu-id="e021e-167">值</span><span class="sxs-lookup"><span data-stu-id="e021e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e021e-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e021e-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e021e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="e021e-169">1.1</span></span>|
|[<span data-ttu-id="e021e-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e021e-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e021e-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e021e-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e021e-172">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="e021e-172">CoercionType: String</span></span>

<span data-ttu-id="e021e-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="e021e-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e021e-174">类型</span><span class="sxs-lookup"><span data-stu-id="e021e-174">Type</span></span>

*   <span data-ttu-id="e021e-175">String</span><span class="sxs-lookup"><span data-stu-id="e021e-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e021e-176">属性：</span><span class="sxs-lookup"><span data-stu-id="e021e-176">Properties:</span></span>

|<span data-ttu-id="e021e-177">姓名</span><span class="sxs-lookup"><span data-stu-id="e021e-177">Name</span></span>| <span data-ttu-id="e021e-178">类型</span><span class="sxs-lookup"><span data-stu-id="e021e-178">Type</span></span>| <span data-ttu-id="e021e-179">说明</span><span class="sxs-lookup"><span data-stu-id="e021e-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e021e-180">String</span><span class="sxs-lookup"><span data-stu-id="e021e-180">String</span></span>|<span data-ttu-id="e021e-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="e021e-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e021e-182">String</span><span class="sxs-lookup"><span data-stu-id="e021e-182">String</span></span>|<span data-ttu-id="e021e-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="e021e-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e021e-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="e021e-184">Requirements</span></span>

|<span data-ttu-id="e021e-185">要求</span><span class="sxs-lookup"><span data-stu-id="e021e-185">Requirement</span></span>| <span data-ttu-id="e021e-186">值</span><span class="sxs-lookup"><span data-stu-id="e021e-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="e021e-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e021e-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e021e-188">1.1</span><span class="sxs-lookup"><span data-stu-id="e021e-188">1.1</span></span>|
|[<span data-ttu-id="e021e-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e021e-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e021e-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e021e-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="e021e-191">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="e021e-191">EventType: String</span></span>

<span data-ttu-id="e021e-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="e021e-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e021e-193">类型</span><span class="sxs-lookup"><span data-stu-id="e021e-193">Type</span></span>

*   <span data-ttu-id="e021e-194">String</span><span class="sxs-lookup"><span data-stu-id="e021e-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e021e-195">属性：</span><span class="sxs-lookup"><span data-stu-id="e021e-195">Properties:</span></span>

| <span data-ttu-id="e021e-196">姓名</span><span class="sxs-lookup"><span data-stu-id="e021e-196">Name</span></span> | <span data-ttu-id="e021e-197">类型</span><span class="sxs-lookup"><span data-stu-id="e021e-197">Type</span></span> | <span data-ttu-id="e021e-198">说明</span><span class="sxs-lookup"><span data-stu-id="e021e-198">Description</span></span> | <span data-ttu-id="e021e-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="e021e-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="e021e-200">String</span><span class="sxs-lookup"><span data-stu-id="e021e-200">String</span></span> | <span data-ttu-id="e021e-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="e021e-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e021e-202">1.7</span><span class="sxs-lookup"><span data-stu-id="e021e-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="e021e-203">String</span><span class="sxs-lookup"><span data-stu-id="e021e-203">String</span></span> | <span data-ttu-id="e021e-204">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="e021e-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e021e-205">1.5</span><span class="sxs-lookup"><span data-stu-id="e021e-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e021e-206">String</span><span class="sxs-lookup"><span data-stu-id="e021e-206">String</span></span> | <span data-ttu-id="e021e-207">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="e021e-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e021e-208">1.7</span><span class="sxs-lookup"><span data-stu-id="e021e-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e021e-209">String</span><span class="sxs-lookup"><span data-stu-id="e021e-209">String</span></span> | <span data-ttu-id="e021e-210">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="e021e-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e021e-211">1.7</span><span class="sxs-lookup"><span data-stu-id="e021e-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e021e-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="e021e-212">Requirements</span></span>

|<span data-ttu-id="e021e-213">要求</span><span class="sxs-lookup"><span data-stu-id="e021e-213">Requirement</span></span>| <span data-ttu-id="e021e-214">值</span><span class="sxs-lookup"><span data-stu-id="e021e-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="e021e-215">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e021e-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e021e-216">1.5</span><span class="sxs-lookup"><span data-stu-id="e021e-216">1.5</span></span> |
|[<span data-ttu-id="e021e-217">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e021e-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e021e-218">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e021e-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e021e-219">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="e021e-219">SourceProperty: String</span></span>

<span data-ttu-id="e021e-220">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="e021e-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e021e-221">类型</span><span class="sxs-lookup"><span data-stu-id="e021e-221">Type</span></span>

*   <span data-ttu-id="e021e-222">String</span><span class="sxs-lookup"><span data-stu-id="e021e-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e021e-223">属性：</span><span class="sxs-lookup"><span data-stu-id="e021e-223">Properties:</span></span>

|<span data-ttu-id="e021e-224">姓名</span><span class="sxs-lookup"><span data-stu-id="e021e-224">Name</span></span>| <span data-ttu-id="e021e-225">类型</span><span class="sxs-lookup"><span data-stu-id="e021e-225">Type</span></span>| <span data-ttu-id="e021e-226">说明</span><span class="sxs-lookup"><span data-stu-id="e021e-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e021e-227">String</span><span class="sxs-lookup"><span data-stu-id="e021e-227">String</span></span>|<span data-ttu-id="e021e-228">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="e021e-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e021e-229">String</span><span class="sxs-lookup"><span data-stu-id="e021e-229">String</span></span>|<span data-ttu-id="e021e-230">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="e021e-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e021e-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="e021e-231">Requirements</span></span>

|<span data-ttu-id="e021e-232">要求</span><span class="sxs-lookup"><span data-stu-id="e021e-232">Requirement</span></span>| <span data-ttu-id="e021e-233">值</span><span class="sxs-lookup"><span data-stu-id="e021e-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="e021e-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e021e-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e021e-235">1.1</span><span class="sxs-lookup"><span data-stu-id="e021e-235">1.1</span></span>|
|[<span data-ttu-id="e021e-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e021e-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e021e-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e021e-237">Compose or Read</span></span>|
