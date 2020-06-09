---
title: Office 命名空间-要求集1。6
description: 使用邮箱 API 要求集1.6 的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: b0d1643727055c6b7ddb4d03c0488b82b24f3fad
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611454"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="cbe7d-103">Office （邮箱要求集1.6）</span><span class="sxs-lookup"><span data-stu-id="cbe7d-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="cbe7d-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cbe7d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="cbe7d-106">Requirements</span></span>

|<span data-ttu-id="cbe7d-107">要求</span><span class="sxs-lookup"><span data-stu-id="cbe7d-107">Requirement</span></span>| <span data-ttu-id="cbe7d-108">值</span><span class="sxs-lookup"><span data-stu-id="cbe7d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbe7d-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cbe7d-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cbe7d-110">1.1</span><span class="sxs-lookup"><span data-stu-id="cbe7d-110">1.1</span></span>|
|[<span data-ttu-id="cbe7d-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cbe7d-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cbe7d-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cbe7d-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cbe7d-113">属性</span><span class="sxs-lookup"><span data-stu-id="cbe7d-113">Properties</span></span>

| <span data-ttu-id="cbe7d-114">属性</span><span class="sxs-lookup"><span data-stu-id="cbe7d-114">Property</span></span> | <span data-ttu-id="cbe7d-115">型号</span><span class="sxs-lookup"><span data-stu-id="cbe7d-115">Modes</span></span> | <span data-ttu-id="cbe7d-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-116">Return type</span></span> | <span data-ttu-id="cbe7d-117">最低</span><span class="sxs-lookup"><span data-stu-id="cbe7d-117">Minimum</span></span><br><span data-ttu-id="cbe7d-118">要求集</span><span class="sxs-lookup"><span data-stu-id="cbe7d-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cbe7d-119">context</span><span class="sxs-lookup"><span data-stu-id="cbe7d-119">context</span></span>](office.context.md) | <span data-ttu-id="cbe7d-120">撰写</span><span class="sxs-lookup"><span data-stu-id="cbe7d-120">Compose</span></span><br><span data-ttu-id="cbe7d-121">Read</span><span class="sxs-lookup"><span data-stu-id="cbe7d-121">Read</span></span> | [<span data-ttu-id="cbe7d-122">Context</span><span class="sxs-lookup"><span data-stu-id="cbe7d-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="cbe7d-123">1.1</span><span class="sxs-lookup"><span data-stu-id="cbe7d-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="cbe7d-124">枚举</span><span class="sxs-lookup"><span data-stu-id="cbe7d-124">Enumerations</span></span>

| <span data-ttu-id="cbe7d-125">枚举</span><span class="sxs-lookup"><span data-stu-id="cbe7d-125">Enumeration</span></span> | <span data-ttu-id="cbe7d-126">型号</span><span class="sxs-lookup"><span data-stu-id="cbe7d-126">Modes</span></span> | <span data-ttu-id="cbe7d-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-127">Return type</span></span> | <span data-ttu-id="cbe7d-128">最低</span><span class="sxs-lookup"><span data-stu-id="cbe7d-128">Minimum</span></span><br><span data-ttu-id="cbe7d-129">要求集</span><span class="sxs-lookup"><span data-stu-id="cbe7d-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cbe7d-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="cbe7d-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="cbe7d-131">撰写</span><span class="sxs-lookup"><span data-stu-id="cbe7d-131">Compose</span></span><br><span data-ttu-id="cbe7d-132">Read</span><span class="sxs-lookup"><span data-stu-id="cbe7d-132">Read</span></span> | <span data-ttu-id="cbe7d-133">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-133">String</span></span> | [<span data-ttu-id="cbe7d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="cbe7d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cbe7d-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="cbe7d-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="cbe7d-136">撰写</span><span class="sxs-lookup"><span data-stu-id="cbe7d-136">Compose</span></span><br><span data-ttu-id="cbe7d-137">Read</span><span class="sxs-lookup"><span data-stu-id="cbe7d-137">Read</span></span> | <span data-ttu-id="cbe7d-138">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-138">String</span></span> | [<span data-ttu-id="cbe7d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cbe7d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cbe7d-140">EventType</span><span class="sxs-lookup"><span data-stu-id="cbe7d-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="cbe7d-141">撰写</span><span class="sxs-lookup"><span data-stu-id="cbe7d-141">Compose</span></span><br><span data-ttu-id="cbe7d-142">Read</span><span class="sxs-lookup"><span data-stu-id="cbe7d-142">Read</span></span> | <span data-ttu-id="cbe7d-143">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-143">String</span></span> | [<span data-ttu-id="cbe7d-144">1.5</span><span class="sxs-lookup"><span data-stu-id="cbe7d-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="cbe7d-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="cbe7d-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="cbe7d-146">撰写</span><span class="sxs-lookup"><span data-stu-id="cbe7d-146">Compose</span></span><br><span data-ttu-id="cbe7d-147">Read</span><span class="sxs-lookup"><span data-stu-id="cbe7d-147">Read</span></span> | <span data-ttu-id="cbe7d-148">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-148">String</span></span> | [<span data-ttu-id="cbe7d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="cbe7d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="cbe7d-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="cbe7d-150">Namespaces</span></span>

<span data-ttu-id="cbe7d-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6)：包含许多特定于 Outlook 的枚举，例如、、、、、 `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` 和 `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="cbe7d-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="cbe7d-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="cbe7d-153">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="cbe7d-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="cbe7d-155">类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-155">Type</span></span>

*   <span data-ttu-id="cbe7d-156">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cbe7d-157">属性：</span><span class="sxs-lookup"><span data-stu-id="cbe7d-157">Properties:</span></span>

|<span data-ttu-id="cbe7d-158">名称</span><span class="sxs-lookup"><span data-stu-id="cbe7d-158">Name</span></span>| <span data-ttu-id="cbe7d-159">类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-159">Type</span></span>| <span data-ttu-id="cbe7d-160">说明</span><span class="sxs-lookup"><span data-stu-id="cbe7d-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="cbe7d-161">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-161">String</span></span>|<span data-ttu-id="cbe7d-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="cbe7d-163">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-163">String</span></span>|<span data-ttu-id="cbe7d-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cbe7d-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="cbe7d-165">Requirements</span></span>

|<span data-ttu-id="cbe7d-166">要求</span><span class="sxs-lookup"><span data-stu-id="cbe7d-166">Requirement</span></span>| <span data-ttu-id="cbe7d-167">值</span><span class="sxs-lookup"><span data-stu-id="cbe7d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbe7d-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cbe7d-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cbe7d-169">1.1</span><span class="sxs-lookup"><span data-stu-id="cbe7d-169">1.1</span></span>|
|[<span data-ttu-id="cbe7d-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cbe7d-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cbe7d-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cbe7d-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="cbe7d-172">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-172">CoercionType: String</span></span>

<span data-ttu-id="cbe7d-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cbe7d-174">类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-174">Type</span></span>

*   <span data-ttu-id="cbe7d-175">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cbe7d-176">属性：</span><span class="sxs-lookup"><span data-stu-id="cbe7d-176">Properties:</span></span>

|<span data-ttu-id="cbe7d-177">名称</span><span class="sxs-lookup"><span data-stu-id="cbe7d-177">Name</span></span>| <span data-ttu-id="cbe7d-178">类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-178">Type</span></span>| <span data-ttu-id="cbe7d-179">说明</span><span class="sxs-lookup"><span data-stu-id="cbe7d-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="cbe7d-180">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-180">String</span></span>|<span data-ttu-id="cbe7d-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="cbe7d-182">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-182">String</span></span>|<span data-ttu-id="cbe7d-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cbe7d-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="cbe7d-184">Requirements</span></span>

|<span data-ttu-id="cbe7d-185">要求</span><span class="sxs-lookup"><span data-stu-id="cbe7d-185">Requirement</span></span>| <span data-ttu-id="cbe7d-186">值</span><span class="sxs-lookup"><span data-stu-id="cbe7d-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbe7d-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cbe7d-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cbe7d-188">1.1</span><span class="sxs-lookup"><span data-stu-id="cbe7d-188">1.1</span></span>|
|[<span data-ttu-id="cbe7d-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cbe7d-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cbe7d-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cbe7d-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="cbe7d-191">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-191">EventType: String</span></span>

<span data-ttu-id="cbe7d-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="cbe7d-193">类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-193">Type</span></span>

*   <span data-ttu-id="cbe7d-194">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cbe7d-195">属性：</span><span class="sxs-lookup"><span data-stu-id="cbe7d-195">Properties:</span></span>

| <span data-ttu-id="cbe7d-196">名称</span><span class="sxs-lookup"><span data-stu-id="cbe7d-196">Name</span></span> | <span data-ttu-id="cbe7d-197">类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-197">Type</span></span> | <span data-ttu-id="cbe7d-198">Description</span><span class="sxs-lookup"><span data-stu-id="cbe7d-198">Description</span></span> | <span data-ttu-id="cbe7d-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="cbe7d-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="cbe7d-200">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-200">String</span></span> | <span data-ttu-id="cbe7d-201">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="cbe7d-202">1.5</span><span class="sxs-lookup"><span data-stu-id="cbe7d-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cbe7d-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="cbe7d-203">Requirements</span></span>

|<span data-ttu-id="cbe7d-204">要求</span><span class="sxs-lookup"><span data-stu-id="cbe7d-204">Requirement</span></span>| <span data-ttu-id="cbe7d-205">值</span><span class="sxs-lookup"><span data-stu-id="cbe7d-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbe7d-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cbe7d-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cbe7d-207">1.5</span><span class="sxs-lookup"><span data-stu-id="cbe7d-207">1.5</span></span> |
|[<span data-ttu-id="cbe7d-208">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cbe7d-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cbe7d-209">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cbe7d-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="cbe7d-210">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-210">SourceProperty: String</span></span>

<span data-ttu-id="cbe7d-211">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cbe7d-212">类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-212">Type</span></span>

*   <span data-ttu-id="cbe7d-213">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cbe7d-214">属性：</span><span class="sxs-lookup"><span data-stu-id="cbe7d-214">Properties:</span></span>

|<span data-ttu-id="cbe7d-215">名称</span><span class="sxs-lookup"><span data-stu-id="cbe7d-215">Name</span></span>| <span data-ttu-id="cbe7d-216">类型</span><span class="sxs-lookup"><span data-stu-id="cbe7d-216">Type</span></span>| <span data-ttu-id="cbe7d-217">说明</span><span class="sxs-lookup"><span data-stu-id="cbe7d-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="cbe7d-218">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-218">String</span></span>|<span data-ttu-id="cbe7d-219">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="cbe7d-220">String</span><span class="sxs-lookup"><span data-stu-id="cbe7d-220">String</span></span>|<span data-ttu-id="cbe7d-221">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="cbe7d-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cbe7d-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="cbe7d-222">Requirements</span></span>

|<span data-ttu-id="cbe7d-223">要求</span><span class="sxs-lookup"><span data-stu-id="cbe7d-223">Requirement</span></span>| <span data-ttu-id="cbe7d-224">值</span><span class="sxs-lookup"><span data-stu-id="cbe7d-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbe7d-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cbe7d-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cbe7d-226">1.1</span><span class="sxs-lookup"><span data-stu-id="cbe7d-226">1.1</span></span>|
|[<span data-ttu-id="cbe7d-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cbe7d-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cbe7d-228">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cbe7d-228">Compose or Read</span></span>|
