---
title: Office 命名空间-要求集1。6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e15f01db9423a9df38608f18098d2c808f5d944b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814666"
---
# <a name="office"></a><span data-ttu-id="8d02a-102">Office</span><span class="sxs-lookup"><span data-stu-id="8d02a-102">Office</span></span>

<span data-ttu-id="8d02a-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="8d02a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8d02a-105">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-105">Requirements</span></span>

|<span data-ttu-id="8d02a-106">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-106">Requirement</span></span>| <span data-ttu-id="8d02a-107">值</span><span class="sxs-lookup"><span data-stu-id="8d02a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d02a-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8d02a-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8d02a-109">1.1</span><span class="sxs-lookup"><span data-stu-id="8d02a-109">1.1</span></span>|
|[<span data-ttu-id="8d02a-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8d02a-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d02a-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8d02a-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8d02a-112">属性</span><span class="sxs-lookup"><span data-stu-id="8d02a-112">Properties</span></span>

| <span data-ttu-id="8d02a-113">属性</span><span class="sxs-lookup"><span data-stu-id="8d02a-113">Property</span></span> | <span data-ttu-id="8d02a-114">型号</span><span class="sxs-lookup"><span data-stu-id="8d02a-114">Modes</span></span> | <span data-ttu-id="8d02a-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-115">Return type</span></span> | <span data-ttu-id="8d02a-116">最低</span><span class="sxs-lookup"><span data-stu-id="8d02a-116">Minimum</span></span><br><span data-ttu-id="8d02a-117">要求集</span><span class="sxs-lookup"><span data-stu-id="8d02a-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8d02a-118">context</span><span class="sxs-lookup"><span data-stu-id="8d02a-118">context</span></span>](office.context.md) | <span data-ttu-id="8d02a-119">撰写</span><span class="sxs-lookup"><span data-stu-id="8d02a-119">Compose</span></span><br><span data-ttu-id="8d02a-120">读取</span><span class="sxs-lookup"><span data-stu-id="8d02a-120">Read</span></span> | [<span data-ttu-id="8d02a-121">Context</span><span class="sxs-lookup"><span data-stu-id="8d02a-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="8d02a-122">1.1</span><span class="sxs-lookup"><span data-stu-id="8d02a-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="8d02a-123">枚举</span><span class="sxs-lookup"><span data-stu-id="8d02a-123">Enumerations</span></span>

| <span data-ttu-id="8d02a-124">枚举</span><span class="sxs-lookup"><span data-stu-id="8d02a-124">Enumeration</span></span> | <span data-ttu-id="8d02a-125">型号</span><span class="sxs-lookup"><span data-stu-id="8d02a-125">Modes</span></span> | <span data-ttu-id="8d02a-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-126">Return type</span></span> | <span data-ttu-id="8d02a-127">最低</span><span class="sxs-lookup"><span data-stu-id="8d02a-127">Minimum</span></span><br><span data-ttu-id="8d02a-128">要求集</span><span class="sxs-lookup"><span data-stu-id="8d02a-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8d02a-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8d02a-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8d02a-130">撰写</span><span class="sxs-lookup"><span data-stu-id="8d02a-130">Compose</span></span><br><span data-ttu-id="8d02a-131">读取</span><span class="sxs-lookup"><span data-stu-id="8d02a-131">Read</span></span> | <span data-ttu-id="8d02a-132">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-132">String</span></span> | [<span data-ttu-id="8d02a-133">1.1</span><span class="sxs-lookup"><span data-stu-id="8d02a-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8d02a-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8d02a-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8d02a-135">撰写</span><span class="sxs-lookup"><span data-stu-id="8d02a-135">Compose</span></span><br><span data-ttu-id="8d02a-136">读取</span><span class="sxs-lookup"><span data-stu-id="8d02a-136">Read</span></span> | <span data-ttu-id="8d02a-137">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-137">String</span></span> | [<span data-ttu-id="8d02a-138">1.1</span><span class="sxs-lookup"><span data-stu-id="8d02a-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8d02a-139">EventType</span><span class="sxs-lookup"><span data-stu-id="8d02a-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8d02a-140">撰写</span><span class="sxs-lookup"><span data-stu-id="8d02a-140">Compose</span></span><br><span data-ttu-id="8d02a-141">读取</span><span class="sxs-lookup"><span data-stu-id="8d02a-141">Read</span></span> | <span data-ttu-id="8d02a-142">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-142">String</span></span> | [<span data-ttu-id="8d02a-143">1.5</span><span class="sxs-lookup"><span data-stu-id="8d02a-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="8d02a-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8d02a-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8d02a-145">撰写</span><span class="sxs-lookup"><span data-stu-id="8d02a-145">Compose</span></span><br><span data-ttu-id="8d02a-146">读取</span><span class="sxs-lookup"><span data-stu-id="8d02a-146">Read</span></span> | <span data-ttu-id="8d02a-147">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-147">String</span></span> | [<span data-ttu-id="8d02a-148">1.1</span><span class="sxs-lookup"><span data-stu-id="8d02a-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="8d02a-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="8d02a-149">Namespaces</span></span>

<span data-ttu-id="8d02a-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="8d02a-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="8d02a-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="8d02a-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="8d02a-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="8d02a-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="8d02a-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="8d02a-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8d02a-154">类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-154">Type</span></span>

*   <span data-ttu-id="8d02a-155">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8d02a-156">属性：</span><span class="sxs-lookup"><span data-stu-id="8d02a-156">Properties:</span></span>

|<span data-ttu-id="8d02a-157">名称</span><span class="sxs-lookup"><span data-stu-id="8d02a-157">Name</span></span>| <span data-ttu-id="8d02a-158">类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-158">Type</span></span>| <span data-ttu-id="8d02a-159">说明</span><span class="sxs-lookup"><span data-stu-id="8d02a-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8d02a-160">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-160">String</span></span>|<span data-ttu-id="8d02a-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="8d02a-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8d02a-162">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-162">String</span></span>|<span data-ttu-id="8d02a-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="8d02a-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d02a-164">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-164">Requirements</span></span>

|<span data-ttu-id="8d02a-165">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-165">Requirement</span></span>| <span data-ttu-id="8d02a-166">值</span><span class="sxs-lookup"><span data-stu-id="8d02a-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d02a-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8d02a-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8d02a-168">1.1</span><span class="sxs-lookup"><span data-stu-id="8d02a-168">1.1</span></span>|
|[<span data-ttu-id="8d02a-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8d02a-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d02a-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8d02a-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="8d02a-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="8d02a-171">CoercionType: String</span></span>

<span data-ttu-id="8d02a-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="8d02a-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8d02a-173">类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-173">Type</span></span>

*   <span data-ttu-id="8d02a-174">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8d02a-175">属性：</span><span class="sxs-lookup"><span data-stu-id="8d02a-175">Properties:</span></span>

|<span data-ttu-id="8d02a-176">名称</span><span class="sxs-lookup"><span data-stu-id="8d02a-176">Name</span></span>| <span data-ttu-id="8d02a-177">类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-177">Type</span></span>| <span data-ttu-id="8d02a-178">说明</span><span class="sxs-lookup"><span data-stu-id="8d02a-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8d02a-179">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-179">String</span></span>|<span data-ttu-id="8d02a-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="8d02a-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8d02a-181">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-181">String</span></span>|<span data-ttu-id="8d02a-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="8d02a-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d02a-183">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-183">Requirements</span></span>

|<span data-ttu-id="8d02a-184">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-184">Requirement</span></span>| <span data-ttu-id="8d02a-185">值</span><span class="sxs-lookup"><span data-stu-id="8d02a-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d02a-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8d02a-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8d02a-187">1.1</span><span class="sxs-lookup"><span data-stu-id="8d02a-187">1.1</span></span>|
|[<span data-ttu-id="8d02a-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8d02a-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d02a-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8d02a-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="8d02a-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="8d02a-190">EventType: String</span></span>

<span data-ttu-id="8d02a-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="8d02a-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8d02a-192">类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-192">Type</span></span>

*   <span data-ttu-id="8d02a-193">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8d02a-194">属性：</span><span class="sxs-lookup"><span data-stu-id="8d02a-194">Properties:</span></span>

| <span data-ttu-id="8d02a-195">名称</span><span class="sxs-lookup"><span data-stu-id="8d02a-195">Name</span></span> | <span data-ttu-id="8d02a-196">类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-196">Type</span></span> | <span data-ttu-id="8d02a-197">说明</span><span class="sxs-lookup"><span data-stu-id="8d02a-197">Description</span></span> | <span data-ttu-id="8d02a-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="8d02a-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="8d02a-199">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-199">String</span></span> | <span data-ttu-id="8d02a-200">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="8d02a-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="8d02a-201">1.5</span><span class="sxs-lookup"><span data-stu-id="8d02a-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8d02a-202">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-202">Requirements</span></span>

|<span data-ttu-id="8d02a-203">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-203">Requirement</span></span>| <span data-ttu-id="8d02a-204">值</span><span class="sxs-lookup"><span data-stu-id="8d02a-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d02a-205">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8d02a-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8d02a-206">1.5</span><span class="sxs-lookup"><span data-stu-id="8d02a-206">1.5</span></span> |
|[<span data-ttu-id="8d02a-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8d02a-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d02a-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8d02a-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="8d02a-209">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="8d02a-209">SourceProperty: String</span></span>

<span data-ttu-id="8d02a-210">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="8d02a-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8d02a-211">类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-211">Type</span></span>

*   <span data-ttu-id="8d02a-212">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8d02a-213">属性：</span><span class="sxs-lookup"><span data-stu-id="8d02a-213">Properties:</span></span>

|<span data-ttu-id="8d02a-214">名称</span><span class="sxs-lookup"><span data-stu-id="8d02a-214">Name</span></span>| <span data-ttu-id="8d02a-215">类型</span><span class="sxs-lookup"><span data-stu-id="8d02a-215">Type</span></span>| <span data-ttu-id="8d02a-216">说明</span><span class="sxs-lookup"><span data-stu-id="8d02a-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8d02a-217">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-217">String</span></span>|<span data-ttu-id="8d02a-218">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="8d02a-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8d02a-219">String</span><span class="sxs-lookup"><span data-stu-id="8d02a-219">String</span></span>|<span data-ttu-id="8d02a-220">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="8d02a-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8d02a-221">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-221">Requirements</span></span>

|<span data-ttu-id="8d02a-222">要求</span><span class="sxs-lookup"><span data-stu-id="8d02a-222">Requirement</span></span>| <span data-ttu-id="8d02a-223">值</span><span class="sxs-lookup"><span data-stu-id="8d02a-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="8d02a-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8d02a-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8d02a-225">1.1</span><span class="sxs-lookup"><span data-stu-id="8d02a-225">1.1</span></span>|
|[<span data-ttu-id="8d02a-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8d02a-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8d02a-227">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8d02a-227">Compose or Read</span></span>|
