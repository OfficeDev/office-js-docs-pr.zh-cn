---
title: Office 命名空间-要求集1。6
description: 使用邮箱 API 要求集1.6 的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: dc7f62cc3f01e56f6c05b6cf40a4b73e87aea5e4
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891311"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="0ede9-103">Office （邮箱要求集1.6）</span><span class="sxs-lookup"><span data-stu-id="0ede9-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="0ede9-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="0ede9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0ede9-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="0ede9-106">Requirements</span></span>

|<span data-ttu-id="0ede9-107">要求</span><span class="sxs-lookup"><span data-stu-id="0ede9-107">Requirement</span></span>| <span data-ttu-id="0ede9-108">值</span><span class="sxs-lookup"><span data-stu-id="0ede9-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ede9-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0ede9-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ede9-110">1.1</span><span class="sxs-lookup"><span data-stu-id="0ede9-110">1.1</span></span>|
|[<span data-ttu-id="0ede9-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ede9-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ede9-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ede9-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="0ede9-113">属性</span><span class="sxs-lookup"><span data-stu-id="0ede9-113">Properties</span></span>

| <span data-ttu-id="0ede9-114">属性</span><span class="sxs-lookup"><span data-stu-id="0ede9-114">Property</span></span> | <span data-ttu-id="0ede9-115">型号</span><span class="sxs-lookup"><span data-stu-id="0ede9-115">Modes</span></span> | <span data-ttu-id="0ede9-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-116">Return type</span></span> | <span data-ttu-id="0ede9-117">最低</span><span class="sxs-lookup"><span data-stu-id="0ede9-117">Minimum</span></span><br><span data-ttu-id="0ede9-118">要求集</span><span class="sxs-lookup"><span data-stu-id="0ede9-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0ede9-119">context</span><span class="sxs-lookup"><span data-stu-id="0ede9-119">context</span></span>](office.context.md) | <span data-ttu-id="0ede9-120">撰写</span><span class="sxs-lookup"><span data-stu-id="0ede9-120">Compose</span></span><br><span data-ttu-id="0ede9-121">读取</span><span class="sxs-lookup"><span data-stu-id="0ede9-121">Read</span></span> | [<span data-ttu-id="0ede9-122">Context</span><span class="sxs-lookup"><span data-stu-id="0ede9-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="0ede9-123">1.1</span><span class="sxs-lookup"><span data-stu-id="0ede9-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="0ede9-124">枚举</span><span class="sxs-lookup"><span data-stu-id="0ede9-124">Enumerations</span></span>

| <span data-ttu-id="0ede9-125">枚举</span><span class="sxs-lookup"><span data-stu-id="0ede9-125">Enumeration</span></span> | <span data-ttu-id="0ede9-126">型号</span><span class="sxs-lookup"><span data-stu-id="0ede9-126">Modes</span></span> | <span data-ttu-id="0ede9-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-127">Return type</span></span> | <span data-ttu-id="0ede9-128">最低</span><span class="sxs-lookup"><span data-stu-id="0ede9-128">Minimum</span></span><br><span data-ttu-id="0ede9-129">要求集</span><span class="sxs-lookup"><span data-stu-id="0ede9-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0ede9-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0ede9-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0ede9-131">撰写</span><span class="sxs-lookup"><span data-stu-id="0ede9-131">Compose</span></span><br><span data-ttu-id="0ede9-132">读取</span><span class="sxs-lookup"><span data-stu-id="0ede9-132">Read</span></span> | <span data-ttu-id="0ede9-133">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-133">String</span></span> | [<span data-ttu-id="0ede9-134">1.1</span><span class="sxs-lookup"><span data-stu-id="0ede9-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0ede9-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0ede9-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0ede9-136">撰写</span><span class="sxs-lookup"><span data-stu-id="0ede9-136">Compose</span></span><br><span data-ttu-id="0ede9-137">读取</span><span class="sxs-lookup"><span data-stu-id="0ede9-137">Read</span></span> | <span data-ttu-id="0ede9-138">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-138">String</span></span> | [<span data-ttu-id="0ede9-139">1.1</span><span class="sxs-lookup"><span data-stu-id="0ede9-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0ede9-140">EventType</span><span class="sxs-lookup"><span data-stu-id="0ede9-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="0ede9-141">撰写</span><span class="sxs-lookup"><span data-stu-id="0ede9-141">Compose</span></span><br><span data-ttu-id="0ede9-142">读取</span><span class="sxs-lookup"><span data-stu-id="0ede9-142">Read</span></span> | <span data-ttu-id="0ede9-143">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-143">String</span></span> | [<span data-ttu-id="0ede9-144">1.5</span><span class="sxs-lookup"><span data-stu-id="0ede9-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="0ede9-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0ede9-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0ede9-146">撰写</span><span class="sxs-lookup"><span data-stu-id="0ede9-146">Compose</span></span><br><span data-ttu-id="0ede9-147">读取</span><span class="sxs-lookup"><span data-stu-id="0ede9-147">Read</span></span> | <span data-ttu-id="0ede9-148">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-148">String</span></span> | [<span data-ttu-id="0ede9-149">1.1</span><span class="sxs-lookup"><span data-stu-id="0ede9-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="0ede9-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="0ede9-150">Namespaces</span></span>

<span data-ttu-id="0ede9-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="0ede9-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="0ede9-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="0ede9-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="0ede9-153">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="0ede9-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="0ede9-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="0ede9-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0ede9-155">类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-155">Type</span></span>

*   <span data-ttu-id="0ede9-156">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0ede9-157">属性：</span><span class="sxs-lookup"><span data-stu-id="0ede9-157">Properties:</span></span>

|<span data-ttu-id="0ede9-158">姓名</span><span class="sxs-lookup"><span data-stu-id="0ede9-158">Name</span></span>| <span data-ttu-id="0ede9-159">类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-159">Type</span></span>| <span data-ttu-id="0ede9-160">说明</span><span class="sxs-lookup"><span data-stu-id="0ede9-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0ede9-161">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-161">String</span></span>|<span data-ttu-id="0ede9-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="0ede9-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0ede9-163">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-163">String</span></span>|<span data-ttu-id="0ede9-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="0ede9-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ede9-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="0ede9-165">Requirements</span></span>

|<span data-ttu-id="0ede9-166">要求</span><span class="sxs-lookup"><span data-stu-id="0ede9-166">Requirement</span></span>| <span data-ttu-id="0ede9-167">值</span><span class="sxs-lookup"><span data-stu-id="0ede9-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ede9-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0ede9-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ede9-169">1.1</span><span class="sxs-lookup"><span data-stu-id="0ede9-169">1.1</span></span>|
|[<span data-ttu-id="0ede9-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ede9-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ede9-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ede9-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="0ede9-172">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="0ede9-172">CoercionType: String</span></span>

<span data-ttu-id="0ede9-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="0ede9-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0ede9-174">类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-174">Type</span></span>

*   <span data-ttu-id="0ede9-175">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0ede9-176">属性：</span><span class="sxs-lookup"><span data-stu-id="0ede9-176">Properties:</span></span>

|<span data-ttu-id="0ede9-177">姓名</span><span class="sxs-lookup"><span data-stu-id="0ede9-177">Name</span></span>| <span data-ttu-id="0ede9-178">类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-178">Type</span></span>| <span data-ttu-id="0ede9-179">说明</span><span class="sxs-lookup"><span data-stu-id="0ede9-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0ede9-180">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-180">String</span></span>|<span data-ttu-id="0ede9-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="0ede9-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0ede9-182">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-182">String</span></span>|<span data-ttu-id="0ede9-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="0ede9-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ede9-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="0ede9-184">Requirements</span></span>

|<span data-ttu-id="0ede9-185">要求</span><span class="sxs-lookup"><span data-stu-id="0ede9-185">Requirement</span></span>| <span data-ttu-id="0ede9-186">值</span><span class="sxs-lookup"><span data-stu-id="0ede9-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ede9-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0ede9-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ede9-188">1.1</span><span class="sxs-lookup"><span data-stu-id="0ede9-188">1.1</span></span>|
|[<span data-ttu-id="0ede9-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ede9-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ede9-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ede9-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="0ede9-191">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="0ede9-191">EventType: String</span></span>

<span data-ttu-id="0ede9-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="0ede9-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="0ede9-193">类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-193">Type</span></span>

*   <span data-ttu-id="0ede9-194">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0ede9-195">属性：</span><span class="sxs-lookup"><span data-stu-id="0ede9-195">Properties:</span></span>

| <span data-ttu-id="0ede9-196">姓名</span><span class="sxs-lookup"><span data-stu-id="0ede9-196">Name</span></span> | <span data-ttu-id="0ede9-197">类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-197">Type</span></span> | <span data-ttu-id="0ede9-198">说明</span><span class="sxs-lookup"><span data-stu-id="0ede9-198">Description</span></span> | <span data-ttu-id="0ede9-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="0ede9-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="0ede9-200">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-200">String</span></span> | <span data-ttu-id="0ede9-201">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="0ede9-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="0ede9-202">1.5</span><span class="sxs-lookup"><span data-stu-id="0ede9-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0ede9-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="0ede9-203">Requirements</span></span>

|<span data-ttu-id="0ede9-204">要求</span><span class="sxs-lookup"><span data-stu-id="0ede9-204">Requirement</span></span>| <span data-ttu-id="0ede9-205">值</span><span class="sxs-lookup"><span data-stu-id="0ede9-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ede9-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0ede9-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ede9-207">1.5</span><span class="sxs-lookup"><span data-stu-id="0ede9-207">1.5</span></span> |
|[<span data-ttu-id="0ede9-208">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ede9-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ede9-209">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ede9-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="0ede9-210">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="0ede9-210">SourceProperty: String</span></span>

<span data-ttu-id="0ede9-211">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="0ede9-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0ede9-212">类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-212">Type</span></span>

*   <span data-ttu-id="0ede9-213">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0ede9-214">属性：</span><span class="sxs-lookup"><span data-stu-id="0ede9-214">Properties:</span></span>

|<span data-ttu-id="0ede9-215">姓名</span><span class="sxs-lookup"><span data-stu-id="0ede9-215">Name</span></span>| <span data-ttu-id="0ede9-216">类型</span><span class="sxs-lookup"><span data-stu-id="0ede9-216">Type</span></span>| <span data-ttu-id="0ede9-217">说明</span><span class="sxs-lookup"><span data-stu-id="0ede9-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0ede9-218">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-218">String</span></span>|<span data-ttu-id="0ede9-219">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="0ede9-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0ede9-220">String</span><span class="sxs-lookup"><span data-stu-id="0ede9-220">String</span></span>|<span data-ttu-id="0ede9-221">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="0ede9-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ede9-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="0ede9-222">Requirements</span></span>

|<span data-ttu-id="0ede9-223">要求</span><span class="sxs-lookup"><span data-stu-id="0ede9-223">Requirement</span></span>| <span data-ttu-id="0ede9-224">值</span><span class="sxs-lookup"><span data-stu-id="0ede9-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ede9-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0ede9-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ede9-226">1.1</span><span class="sxs-lookup"><span data-stu-id="0ede9-226">1.1</span></span>|
|[<span data-ttu-id="0ede9-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0ede9-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ede9-228">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0ede9-228">Compose or Read</span></span>|
