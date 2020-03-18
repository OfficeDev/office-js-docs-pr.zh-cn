---
title: Office 命名空间-要求集1。6
description: Outlook 外接程序 API 的顶级命名空间的对象模型（邮箱 API 1.6 版本）。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ae2f863e054016636ebffc3ff3925cee018036a1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717647"
---
# <a name="office"></a><span data-ttu-id="d09b9-103">Office</span><span class="sxs-lookup"><span data-stu-id="d09b9-103">Office</span></span>

<span data-ttu-id="d09b9-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="d09b9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d09b9-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="d09b9-106">Requirements</span></span>

|<span data-ttu-id="d09b9-107">要求</span><span class="sxs-lookup"><span data-stu-id="d09b9-107">Requirement</span></span>| <span data-ttu-id="d09b9-108">值</span><span class="sxs-lookup"><span data-stu-id="d09b9-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d09b9-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d09b9-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d09b9-110">1.1</span><span class="sxs-lookup"><span data-stu-id="d09b9-110">1.1</span></span>|
|[<span data-ttu-id="d09b9-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d09b9-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d09b9-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d09b9-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d09b9-113">属性</span><span class="sxs-lookup"><span data-stu-id="d09b9-113">Properties</span></span>

| <span data-ttu-id="d09b9-114">属性</span><span class="sxs-lookup"><span data-stu-id="d09b9-114">Property</span></span> | <span data-ttu-id="d09b9-115">型号</span><span class="sxs-lookup"><span data-stu-id="d09b9-115">Modes</span></span> | <span data-ttu-id="d09b9-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-116">Return type</span></span> | <span data-ttu-id="d09b9-117">最低</span><span class="sxs-lookup"><span data-stu-id="d09b9-117">Minimum</span></span><br><span data-ttu-id="d09b9-118">要求集</span><span class="sxs-lookup"><span data-stu-id="d09b9-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d09b9-119">context</span><span class="sxs-lookup"><span data-stu-id="d09b9-119">context</span></span>](office.context.md) | <span data-ttu-id="d09b9-120">撰写</span><span class="sxs-lookup"><span data-stu-id="d09b9-120">Compose</span></span><br><span data-ttu-id="d09b9-121">读取</span><span class="sxs-lookup"><span data-stu-id="d09b9-121">Read</span></span> | [<span data-ttu-id="d09b9-122">Context</span><span class="sxs-lookup"><span data-stu-id="d09b9-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="d09b9-123">1.1</span><span class="sxs-lookup"><span data-stu-id="d09b9-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="d09b9-124">枚举</span><span class="sxs-lookup"><span data-stu-id="d09b9-124">Enumerations</span></span>

| <span data-ttu-id="d09b9-125">枚举</span><span class="sxs-lookup"><span data-stu-id="d09b9-125">Enumeration</span></span> | <span data-ttu-id="d09b9-126">型号</span><span class="sxs-lookup"><span data-stu-id="d09b9-126">Modes</span></span> | <span data-ttu-id="d09b9-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-127">Return type</span></span> | <span data-ttu-id="d09b9-128">最低</span><span class="sxs-lookup"><span data-stu-id="d09b9-128">Minimum</span></span><br><span data-ttu-id="d09b9-129">要求集</span><span class="sxs-lookup"><span data-stu-id="d09b9-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d09b9-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d09b9-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d09b9-131">撰写</span><span class="sxs-lookup"><span data-stu-id="d09b9-131">Compose</span></span><br><span data-ttu-id="d09b9-132">读取</span><span class="sxs-lookup"><span data-stu-id="d09b9-132">Read</span></span> | <span data-ttu-id="d09b9-133">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-133">String</span></span> | [<span data-ttu-id="d09b9-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d09b9-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d09b9-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d09b9-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d09b9-136">撰写</span><span class="sxs-lookup"><span data-stu-id="d09b9-136">Compose</span></span><br><span data-ttu-id="d09b9-137">读取</span><span class="sxs-lookup"><span data-stu-id="d09b9-137">Read</span></span> | <span data-ttu-id="d09b9-138">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-138">String</span></span> | [<span data-ttu-id="d09b9-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d09b9-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d09b9-140">EventType</span><span class="sxs-lookup"><span data-stu-id="d09b9-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d09b9-141">撰写</span><span class="sxs-lookup"><span data-stu-id="d09b9-141">Compose</span></span><br><span data-ttu-id="d09b9-142">读取</span><span class="sxs-lookup"><span data-stu-id="d09b9-142">Read</span></span> | <span data-ttu-id="d09b9-143">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-143">String</span></span> | [<span data-ttu-id="d09b9-144">1.5</span><span class="sxs-lookup"><span data-stu-id="d09b9-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="d09b9-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d09b9-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d09b9-146">撰写</span><span class="sxs-lookup"><span data-stu-id="d09b9-146">Compose</span></span><br><span data-ttu-id="d09b9-147">读取</span><span class="sxs-lookup"><span data-stu-id="d09b9-147">Read</span></span> | <span data-ttu-id="d09b9-148">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-148">String</span></span> | [<span data-ttu-id="d09b9-149">1.1</span><span class="sxs-lookup"><span data-stu-id="d09b9-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="d09b9-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="d09b9-150">Namespaces</span></span>

<span data-ttu-id="d09b9-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="d09b9-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="d09b9-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="d09b9-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d09b9-153">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="d09b9-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="d09b9-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="d09b9-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d09b9-155">类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-155">Type</span></span>

*   <span data-ttu-id="d09b9-156">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d09b9-157">属性：</span><span class="sxs-lookup"><span data-stu-id="d09b9-157">Properties:</span></span>

|<span data-ttu-id="d09b9-158">姓名</span><span class="sxs-lookup"><span data-stu-id="d09b9-158">Name</span></span>| <span data-ttu-id="d09b9-159">类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-159">Type</span></span>| <span data-ttu-id="d09b9-160">说明</span><span class="sxs-lookup"><span data-stu-id="d09b9-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d09b9-161">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-161">String</span></span>|<span data-ttu-id="d09b9-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="d09b9-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d09b9-163">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-163">String</span></span>|<span data-ttu-id="d09b9-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="d09b9-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d09b9-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="d09b9-165">Requirements</span></span>

|<span data-ttu-id="d09b9-166">要求</span><span class="sxs-lookup"><span data-stu-id="d09b9-166">Requirement</span></span>| <span data-ttu-id="d09b9-167">值</span><span class="sxs-lookup"><span data-stu-id="d09b9-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d09b9-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d09b9-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d09b9-169">1.1</span><span class="sxs-lookup"><span data-stu-id="d09b9-169">1.1</span></span>|
|[<span data-ttu-id="d09b9-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d09b9-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d09b9-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d09b9-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d09b9-172">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="d09b9-172">CoercionType: String</span></span>

<span data-ttu-id="d09b9-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="d09b9-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d09b9-174">类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-174">Type</span></span>

*   <span data-ttu-id="d09b9-175">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d09b9-176">属性：</span><span class="sxs-lookup"><span data-stu-id="d09b9-176">Properties:</span></span>

|<span data-ttu-id="d09b9-177">姓名</span><span class="sxs-lookup"><span data-stu-id="d09b9-177">Name</span></span>| <span data-ttu-id="d09b9-178">类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-178">Type</span></span>| <span data-ttu-id="d09b9-179">说明</span><span class="sxs-lookup"><span data-stu-id="d09b9-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d09b9-180">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-180">String</span></span>|<span data-ttu-id="d09b9-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d09b9-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d09b9-182">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-182">String</span></span>|<span data-ttu-id="d09b9-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d09b9-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d09b9-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="d09b9-184">Requirements</span></span>

|<span data-ttu-id="d09b9-185">要求</span><span class="sxs-lookup"><span data-stu-id="d09b9-185">Requirement</span></span>| <span data-ttu-id="d09b9-186">值</span><span class="sxs-lookup"><span data-stu-id="d09b9-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="d09b9-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d09b9-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d09b9-188">1.1</span><span class="sxs-lookup"><span data-stu-id="d09b9-188">1.1</span></span>|
|[<span data-ttu-id="d09b9-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d09b9-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d09b9-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d09b9-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="d09b9-191">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="d09b9-191">EventType: String</span></span>

<span data-ttu-id="d09b9-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="d09b9-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d09b9-193">类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-193">Type</span></span>

*   <span data-ttu-id="d09b9-194">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d09b9-195">属性：</span><span class="sxs-lookup"><span data-stu-id="d09b9-195">Properties:</span></span>

| <span data-ttu-id="d09b9-196">姓名</span><span class="sxs-lookup"><span data-stu-id="d09b9-196">Name</span></span> | <span data-ttu-id="d09b9-197">类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-197">Type</span></span> | <span data-ttu-id="d09b9-198">说明</span><span class="sxs-lookup"><span data-stu-id="d09b9-198">Description</span></span> | <span data-ttu-id="d09b9-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="d09b9-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="d09b9-200">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-200">String</span></span> | <span data-ttu-id="d09b9-201">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="d09b9-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="d09b9-202">1.5</span><span class="sxs-lookup"><span data-stu-id="d09b9-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d09b9-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="d09b9-203">Requirements</span></span>

|<span data-ttu-id="d09b9-204">要求</span><span class="sxs-lookup"><span data-stu-id="d09b9-204">Requirement</span></span>| <span data-ttu-id="d09b9-205">值</span><span class="sxs-lookup"><span data-stu-id="d09b9-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="d09b9-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d09b9-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d09b9-207">1.5</span><span class="sxs-lookup"><span data-stu-id="d09b9-207">1.5</span></span> |
|[<span data-ttu-id="d09b9-208">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d09b9-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d09b9-209">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d09b9-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d09b9-210">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="d09b9-210">SourceProperty: String</span></span>

<span data-ttu-id="d09b9-211">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="d09b9-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d09b9-212">类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-212">Type</span></span>

*   <span data-ttu-id="d09b9-213">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d09b9-214">属性：</span><span class="sxs-lookup"><span data-stu-id="d09b9-214">Properties:</span></span>

|<span data-ttu-id="d09b9-215">姓名</span><span class="sxs-lookup"><span data-stu-id="d09b9-215">Name</span></span>| <span data-ttu-id="d09b9-216">类型</span><span class="sxs-lookup"><span data-stu-id="d09b9-216">Type</span></span>| <span data-ttu-id="d09b9-217">说明</span><span class="sxs-lookup"><span data-stu-id="d09b9-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d09b9-218">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-218">String</span></span>|<span data-ttu-id="d09b9-219">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="d09b9-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d09b9-220">String</span><span class="sxs-lookup"><span data-stu-id="d09b9-220">String</span></span>|<span data-ttu-id="d09b9-221">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="d09b9-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d09b9-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="d09b9-222">Requirements</span></span>

|<span data-ttu-id="d09b9-223">要求</span><span class="sxs-lookup"><span data-stu-id="d09b9-223">Requirement</span></span>| <span data-ttu-id="d09b9-224">值</span><span class="sxs-lookup"><span data-stu-id="d09b9-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="d09b9-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d09b9-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d09b9-226">1.1</span><span class="sxs-lookup"><span data-stu-id="d09b9-226">1.1</span></span>|
|[<span data-ttu-id="d09b9-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d09b9-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d09b9-228">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d09b9-228">Compose or Read</span></span>|
