---
title: Office 命名空间-要求集1。5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7cc8e6acc60c28b44ec7a2b91bb5e388b2618a31
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554720"
---
# <a name="office"></a><span data-ttu-id="2a761-102">Office</span><span class="sxs-lookup"><span data-stu-id="2a761-102">Office</span></span>

<span data-ttu-id="2a761-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="2a761-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a761-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a761-105">Requirements</span></span>

|<span data-ttu-id="2a761-106">要求</span><span class="sxs-lookup"><span data-stu-id="2a761-106">Requirement</span></span>| <span data-ttu-id="2a761-107">值</span><span class="sxs-lookup"><span data-stu-id="2a761-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a761-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a761-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a761-109">1.1</span><span class="sxs-lookup"><span data-stu-id="2a761-109">1.1</span></span>|
|[<span data-ttu-id="2a761-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a761-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a761-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a761-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2a761-112">属性</span><span class="sxs-lookup"><span data-stu-id="2a761-112">Properties</span></span>

| <span data-ttu-id="2a761-113">属性</span><span class="sxs-lookup"><span data-stu-id="2a761-113">Property</span></span> | <span data-ttu-id="2a761-114">型号</span><span class="sxs-lookup"><span data-stu-id="2a761-114">Modes</span></span> | <span data-ttu-id="2a761-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="2a761-115">Return type</span></span> | <span data-ttu-id="2a761-116">最低</span><span class="sxs-lookup"><span data-stu-id="2a761-116">Minimum</span></span><br><span data-ttu-id="2a761-117">要求集</span><span class="sxs-lookup"><span data-stu-id="2a761-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2a761-118">context</span><span class="sxs-lookup"><span data-stu-id="2a761-118">context</span></span>](office.context.md) | <span data-ttu-id="2a761-119">撰写</span><span class="sxs-lookup"><span data-stu-id="2a761-119">Compose</span></span><br><span data-ttu-id="2a761-120">读取</span><span class="sxs-lookup"><span data-stu-id="2a761-120">Read</span></span> | [<span data-ttu-id="2a761-121">Context</span><span class="sxs-lookup"><span data-stu-id="2a761-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="2a761-122">1.1</span><span class="sxs-lookup"><span data-stu-id="2a761-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2a761-123">枚举</span><span class="sxs-lookup"><span data-stu-id="2a761-123">Enumerations</span></span>

| <span data-ttu-id="2a761-124">枚举</span><span class="sxs-lookup"><span data-stu-id="2a761-124">Enumeration</span></span> | <span data-ttu-id="2a761-125">型号</span><span class="sxs-lookup"><span data-stu-id="2a761-125">Modes</span></span> | <span data-ttu-id="2a761-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="2a761-126">Return type</span></span> | <span data-ttu-id="2a761-127">最低</span><span class="sxs-lookup"><span data-stu-id="2a761-127">Minimum</span></span><br><span data-ttu-id="2a761-128">要求集</span><span class="sxs-lookup"><span data-stu-id="2a761-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2a761-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2a761-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2a761-130">撰写</span><span class="sxs-lookup"><span data-stu-id="2a761-130">Compose</span></span><br><span data-ttu-id="2a761-131">读取</span><span class="sxs-lookup"><span data-stu-id="2a761-131">Read</span></span> | <span data-ttu-id="2a761-132">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-132">String</span></span> | [<span data-ttu-id="2a761-133">1.1</span><span class="sxs-lookup"><span data-stu-id="2a761-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2a761-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2a761-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2a761-135">撰写</span><span class="sxs-lookup"><span data-stu-id="2a761-135">Compose</span></span><br><span data-ttu-id="2a761-136">读取</span><span class="sxs-lookup"><span data-stu-id="2a761-136">Read</span></span> | <span data-ttu-id="2a761-137">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-137">String</span></span> | [<span data-ttu-id="2a761-138">1.1</span><span class="sxs-lookup"><span data-stu-id="2a761-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2a761-139">EventType</span><span class="sxs-lookup"><span data-stu-id="2a761-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2a761-140">撰写</span><span class="sxs-lookup"><span data-stu-id="2a761-140">Compose</span></span><br><span data-ttu-id="2a761-141">读取</span><span class="sxs-lookup"><span data-stu-id="2a761-141">Read</span></span> | <span data-ttu-id="2a761-142">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-142">String</span></span> | [<span data-ttu-id="2a761-143">1.5</span><span class="sxs-lookup"><span data-stu-id="2a761-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2a761-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2a761-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2a761-145">撰写</span><span class="sxs-lookup"><span data-stu-id="2a761-145">Compose</span></span><br><span data-ttu-id="2a761-146">读取</span><span class="sxs-lookup"><span data-stu-id="2a761-146">Read</span></span> | <span data-ttu-id="2a761-147">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-147">String</span></span> | [<span data-ttu-id="2a761-148">1.1</span><span class="sxs-lookup"><span data-stu-id="2a761-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2a761-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="2a761-149">Namespaces</span></span>

<span data-ttu-id="2a761-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="2a761-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2a761-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="2a761-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2a761-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="2a761-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="2a761-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="2a761-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2a761-154">类型</span><span class="sxs-lookup"><span data-stu-id="2a761-154">Type</span></span>

*   <span data-ttu-id="2a761-155">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a761-156">属性：</span><span class="sxs-lookup"><span data-stu-id="2a761-156">Properties:</span></span>

|<span data-ttu-id="2a761-157">名称</span><span class="sxs-lookup"><span data-stu-id="2a761-157">Name</span></span>| <span data-ttu-id="2a761-158">类型</span><span class="sxs-lookup"><span data-stu-id="2a761-158">Type</span></span>| <span data-ttu-id="2a761-159">说明</span><span class="sxs-lookup"><span data-stu-id="2a761-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2a761-160">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-160">String</span></span>|<span data-ttu-id="2a761-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="2a761-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2a761-162">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-162">String</span></span>|<span data-ttu-id="2a761-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="2a761-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a761-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a761-164">Requirements</span></span>

|<span data-ttu-id="2a761-165">要求</span><span class="sxs-lookup"><span data-stu-id="2a761-165">Requirement</span></span>| <span data-ttu-id="2a761-166">值</span><span class="sxs-lookup"><span data-stu-id="2a761-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a761-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a761-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a761-168">1.1</span><span class="sxs-lookup"><span data-stu-id="2a761-168">1.1</span></span>|
|[<span data-ttu-id="2a761-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a761-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a761-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a761-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2a761-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="2a761-171">CoercionType: String</span></span>

<span data-ttu-id="2a761-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="2a761-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2a761-173">类型</span><span class="sxs-lookup"><span data-stu-id="2a761-173">Type</span></span>

*   <span data-ttu-id="2a761-174">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a761-175">属性：</span><span class="sxs-lookup"><span data-stu-id="2a761-175">Properties:</span></span>

|<span data-ttu-id="2a761-176">名称</span><span class="sxs-lookup"><span data-stu-id="2a761-176">Name</span></span>| <span data-ttu-id="2a761-177">类型</span><span class="sxs-lookup"><span data-stu-id="2a761-177">Type</span></span>| <span data-ttu-id="2a761-178">说明</span><span class="sxs-lookup"><span data-stu-id="2a761-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2a761-179">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-179">String</span></span>|<span data-ttu-id="2a761-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="2a761-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2a761-181">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-181">String</span></span>|<span data-ttu-id="2a761-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="2a761-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a761-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a761-183">Requirements</span></span>

|<span data-ttu-id="2a761-184">要求</span><span class="sxs-lookup"><span data-stu-id="2a761-184">Requirement</span></span>| <span data-ttu-id="2a761-185">值</span><span class="sxs-lookup"><span data-stu-id="2a761-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a761-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a761-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a761-187">1.1</span><span class="sxs-lookup"><span data-stu-id="2a761-187">1.1</span></span>|
|[<span data-ttu-id="2a761-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a761-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a761-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a761-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2a761-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="2a761-190">EventType: String</span></span>

<span data-ttu-id="2a761-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="2a761-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2a761-192">类型</span><span class="sxs-lookup"><span data-stu-id="2a761-192">Type</span></span>

*   <span data-ttu-id="2a761-193">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a761-194">属性：</span><span class="sxs-lookup"><span data-stu-id="2a761-194">Properties:</span></span>

| <span data-ttu-id="2a761-195">名称</span><span class="sxs-lookup"><span data-stu-id="2a761-195">Name</span></span> | <span data-ttu-id="2a761-196">类型</span><span class="sxs-lookup"><span data-stu-id="2a761-196">Type</span></span> | <span data-ttu-id="2a761-197">说明</span><span class="sxs-lookup"><span data-stu-id="2a761-197">Description</span></span> | <span data-ttu-id="2a761-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="2a761-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="2a761-199">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-199">String</span></span> | <span data-ttu-id="2a761-200">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="2a761-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2a761-201">1.5</span><span class="sxs-lookup"><span data-stu-id="2a761-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2a761-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a761-202">Requirements</span></span>

|<span data-ttu-id="2a761-203">要求</span><span class="sxs-lookup"><span data-stu-id="2a761-203">Requirement</span></span>| <span data-ttu-id="2a761-204">值</span><span class="sxs-lookup"><span data-stu-id="2a761-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a761-205">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a761-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a761-206">1.5</span><span class="sxs-lookup"><span data-stu-id="2a761-206">1.5</span></span> |
|[<span data-ttu-id="2a761-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a761-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a761-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a761-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2a761-209">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="2a761-209">SourceProperty: String</span></span>

<span data-ttu-id="2a761-210">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="2a761-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2a761-211">类型</span><span class="sxs-lookup"><span data-stu-id="2a761-211">Type</span></span>

*   <span data-ttu-id="2a761-212">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a761-213">属性：</span><span class="sxs-lookup"><span data-stu-id="2a761-213">Properties:</span></span>

|<span data-ttu-id="2a761-214">名称</span><span class="sxs-lookup"><span data-stu-id="2a761-214">Name</span></span>| <span data-ttu-id="2a761-215">类型</span><span class="sxs-lookup"><span data-stu-id="2a761-215">Type</span></span>| <span data-ttu-id="2a761-216">说明</span><span class="sxs-lookup"><span data-stu-id="2a761-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2a761-217">字符串</span><span class="sxs-lookup"><span data-stu-id="2a761-217">String</span></span>|<span data-ttu-id="2a761-218">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="2a761-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2a761-219">String</span><span class="sxs-lookup"><span data-stu-id="2a761-219">String</span></span>|<span data-ttu-id="2a761-220">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="2a761-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a761-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="2a761-221">Requirements</span></span>

|<span data-ttu-id="2a761-222">要求</span><span class="sxs-lookup"><span data-stu-id="2a761-222">Requirement</span></span>| <span data-ttu-id="2a761-223">值</span><span class="sxs-lookup"><span data-stu-id="2a761-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a761-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2a761-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a761-225">1.1</span><span class="sxs-lookup"><span data-stu-id="2a761-225">1.1</span></span>|
|[<span data-ttu-id="2a761-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2a761-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2a761-227">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2a761-227">Compose or Read</span></span>|
