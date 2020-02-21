---
title: Office 命名空间-要求集1。6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0a6360ff7f4e397b878d9a3f744bdbe58347c558
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163659"
---
# <a name="office"></a><span data-ttu-id="96af9-102">Office</span><span class="sxs-lookup"><span data-stu-id="96af9-102">Office</span></span>

<span data-ttu-id="96af9-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="96af9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="96af9-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="96af9-105">Requirements</span></span>

|<span data-ttu-id="96af9-106">要求</span><span class="sxs-lookup"><span data-stu-id="96af9-106">Requirement</span></span>| <span data-ttu-id="96af9-107">值</span><span class="sxs-lookup"><span data-stu-id="96af9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="96af9-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96af9-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96af9-109">1.1</span><span class="sxs-lookup"><span data-stu-id="96af9-109">1.1</span></span>|
|[<span data-ttu-id="96af9-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96af9-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96af9-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96af9-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="96af9-112">属性</span><span class="sxs-lookup"><span data-stu-id="96af9-112">Properties</span></span>

| <span data-ttu-id="96af9-113">属性</span><span class="sxs-lookup"><span data-stu-id="96af9-113">Property</span></span> | <span data-ttu-id="96af9-114">型号</span><span class="sxs-lookup"><span data-stu-id="96af9-114">Modes</span></span> | <span data-ttu-id="96af9-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="96af9-115">Return type</span></span> | <span data-ttu-id="96af9-116">最低</span><span class="sxs-lookup"><span data-stu-id="96af9-116">Minimum</span></span><br><span data-ttu-id="96af9-117">要求集</span><span class="sxs-lookup"><span data-stu-id="96af9-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="96af9-118">context</span><span class="sxs-lookup"><span data-stu-id="96af9-118">context</span></span>](office.context.md) | <span data-ttu-id="96af9-119">撰写</span><span class="sxs-lookup"><span data-stu-id="96af9-119">Compose</span></span><br><span data-ttu-id="96af9-120">读取</span><span class="sxs-lookup"><span data-stu-id="96af9-120">Read</span></span> | [<span data-ttu-id="96af9-121">Context</span><span class="sxs-lookup"><span data-stu-id="96af9-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="96af9-122">1.1</span><span class="sxs-lookup"><span data-stu-id="96af9-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="96af9-123">枚举</span><span class="sxs-lookup"><span data-stu-id="96af9-123">Enumerations</span></span>

| <span data-ttu-id="96af9-124">枚举</span><span class="sxs-lookup"><span data-stu-id="96af9-124">Enumeration</span></span> | <span data-ttu-id="96af9-125">型号</span><span class="sxs-lookup"><span data-stu-id="96af9-125">Modes</span></span> | <span data-ttu-id="96af9-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="96af9-126">Return type</span></span> | <span data-ttu-id="96af9-127">最低</span><span class="sxs-lookup"><span data-stu-id="96af9-127">Minimum</span></span><br><span data-ttu-id="96af9-128">要求集</span><span class="sxs-lookup"><span data-stu-id="96af9-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="96af9-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="96af9-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="96af9-130">撰写</span><span class="sxs-lookup"><span data-stu-id="96af9-130">Compose</span></span><br><span data-ttu-id="96af9-131">读取</span><span class="sxs-lookup"><span data-stu-id="96af9-131">Read</span></span> | <span data-ttu-id="96af9-132">String</span><span class="sxs-lookup"><span data-stu-id="96af9-132">String</span></span> | [<span data-ttu-id="96af9-133">1.1</span><span class="sxs-lookup"><span data-stu-id="96af9-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96af9-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="96af9-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="96af9-135">撰写</span><span class="sxs-lookup"><span data-stu-id="96af9-135">Compose</span></span><br><span data-ttu-id="96af9-136">读取</span><span class="sxs-lookup"><span data-stu-id="96af9-136">Read</span></span> | <span data-ttu-id="96af9-137">String</span><span class="sxs-lookup"><span data-stu-id="96af9-137">String</span></span> | [<span data-ttu-id="96af9-138">1.1</span><span class="sxs-lookup"><span data-stu-id="96af9-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96af9-139">EventType</span><span class="sxs-lookup"><span data-stu-id="96af9-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="96af9-140">撰写</span><span class="sxs-lookup"><span data-stu-id="96af9-140">Compose</span></span><br><span data-ttu-id="96af9-141">读取</span><span class="sxs-lookup"><span data-stu-id="96af9-141">Read</span></span> | <span data-ttu-id="96af9-142">String</span><span class="sxs-lookup"><span data-stu-id="96af9-142">String</span></span> | [<span data-ttu-id="96af9-143">1.5</span><span class="sxs-lookup"><span data-stu-id="96af9-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="96af9-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="96af9-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="96af9-145">撰写</span><span class="sxs-lookup"><span data-stu-id="96af9-145">Compose</span></span><br><span data-ttu-id="96af9-146">读取</span><span class="sxs-lookup"><span data-stu-id="96af9-146">Read</span></span> | <span data-ttu-id="96af9-147">String</span><span class="sxs-lookup"><span data-stu-id="96af9-147">String</span></span> | [<span data-ttu-id="96af9-148">1.1</span><span class="sxs-lookup"><span data-stu-id="96af9-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="96af9-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="96af9-149">Namespaces</span></span>

<span data-ttu-id="96af9-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="96af9-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="96af9-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="96af9-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="96af9-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="96af9-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="96af9-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="96af9-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="96af9-154">类型</span><span class="sxs-lookup"><span data-stu-id="96af9-154">Type</span></span>

*   <span data-ttu-id="96af9-155">String</span><span class="sxs-lookup"><span data-stu-id="96af9-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="96af9-156">属性：</span><span class="sxs-lookup"><span data-stu-id="96af9-156">Properties:</span></span>

|<span data-ttu-id="96af9-157">名称</span><span class="sxs-lookup"><span data-stu-id="96af9-157">Name</span></span>| <span data-ttu-id="96af9-158">类型</span><span class="sxs-lookup"><span data-stu-id="96af9-158">Type</span></span>| <span data-ttu-id="96af9-159">说明</span><span class="sxs-lookup"><span data-stu-id="96af9-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="96af9-160">String</span><span class="sxs-lookup"><span data-stu-id="96af9-160">String</span></span>|<span data-ttu-id="96af9-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="96af9-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="96af9-162">String</span><span class="sxs-lookup"><span data-stu-id="96af9-162">String</span></span>|<span data-ttu-id="96af9-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="96af9-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96af9-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="96af9-164">Requirements</span></span>

|<span data-ttu-id="96af9-165">要求</span><span class="sxs-lookup"><span data-stu-id="96af9-165">Requirement</span></span>| <span data-ttu-id="96af9-166">值</span><span class="sxs-lookup"><span data-stu-id="96af9-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="96af9-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96af9-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96af9-168">1.1</span><span class="sxs-lookup"><span data-stu-id="96af9-168">1.1</span></span>|
|[<span data-ttu-id="96af9-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96af9-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96af9-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96af9-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="96af9-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="96af9-171">CoercionType: String</span></span>

<span data-ttu-id="96af9-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="96af9-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="96af9-173">类型</span><span class="sxs-lookup"><span data-stu-id="96af9-173">Type</span></span>

*   <span data-ttu-id="96af9-174">String</span><span class="sxs-lookup"><span data-stu-id="96af9-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="96af9-175">属性：</span><span class="sxs-lookup"><span data-stu-id="96af9-175">Properties:</span></span>

|<span data-ttu-id="96af9-176">名称</span><span class="sxs-lookup"><span data-stu-id="96af9-176">Name</span></span>| <span data-ttu-id="96af9-177">类型</span><span class="sxs-lookup"><span data-stu-id="96af9-177">Type</span></span>| <span data-ttu-id="96af9-178">说明</span><span class="sxs-lookup"><span data-stu-id="96af9-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="96af9-179">String</span><span class="sxs-lookup"><span data-stu-id="96af9-179">String</span></span>|<span data-ttu-id="96af9-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="96af9-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="96af9-181">String</span><span class="sxs-lookup"><span data-stu-id="96af9-181">String</span></span>|<span data-ttu-id="96af9-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="96af9-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96af9-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="96af9-183">Requirements</span></span>

|<span data-ttu-id="96af9-184">要求</span><span class="sxs-lookup"><span data-stu-id="96af9-184">Requirement</span></span>| <span data-ttu-id="96af9-185">值</span><span class="sxs-lookup"><span data-stu-id="96af9-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="96af9-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96af9-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96af9-187">1.1</span><span class="sxs-lookup"><span data-stu-id="96af9-187">1.1</span></span>|
|[<span data-ttu-id="96af9-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96af9-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96af9-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96af9-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="96af9-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="96af9-190">EventType: String</span></span>

<span data-ttu-id="96af9-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="96af9-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="96af9-192">类型</span><span class="sxs-lookup"><span data-stu-id="96af9-192">Type</span></span>

*   <span data-ttu-id="96af9-193">String</span><span class="sxs-lookup"><span data-stu-id="96af9-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="96af9-194">属性：</span><span class="sxs-lookup"><span data-stu-id="96af9-194">Properties:</span></span>

| <span data-ttu-id="96af9-195">名称</span><span class="sxs-lookup"><span data-stu-id="96af9-195">Name</span></span> | <span data-ttu-id="96af9-196">类型</span><span class="sxs-lookup"><span data-stu-id="96af9-196">Type</span></span> | <span data-ttu-id="96af9-197">说明</span><span class="sxs-lookup"><span data-stu-id="96af9-197">Description</span></span> | <span data-ttu-id="96af9-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="96af9-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="96af9-199">String</span><span class="sxs-lookup"><span data-stu-id="96af9-199">String</span></span> | <span data-ttu-id="96af9-200">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="96af9-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="96af9-201">1.5</span><span class="sxs-lookup"><span data-stu-id="96af9-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="96af9-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="96af9-202">Requirements</span></span>

|<span data-ttu-id="96af9-203">要求</span><span class="sxs-lookup"><span data-stu-id="96af9-203">Requirement</span></span>| <span data-ttu-id="96af9-204">值</span><span class="sxs-lookup"><span data-stu-id="96af9-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="96af9-205">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96af9-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96af9-206">1.5</span><span class="sxs-lookup"><span data-stu-id="96af9-206">1.5</span></span> |
|[<span data-ttu-id="96af9-207">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96af9-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96af9-208">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96af9-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="96af9-209">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="96af9-209">SourceProperty: String</span></span>

<span data-ttu-id="96af9-210">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="96af9-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="96af9-211">类型</span><span class="sxs-lookup"><span data-stu-id="96af9-211">Type</span></span>

*   <span data-ttu-id="96af9-212">String</span><span class="sxs-lookup"><span data-stu-id="96af9-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="96af9-213">属性：</span><span class="sxs-lookup"><span data-stu-id="96af9-213">Properties:</span></span>

|<span data-ttu-id="96af9-214">名称</span><span class="sxs-lookup"><span data-stu-id="96af9-214">Name</span></span>| <span data-ttu-id="96af9-215">类型</span><span class="sxs-lookup"><span data-stu-id="96af9-215">Type</span></span>| <span data-ttu-id="96af9-216">说明</span><span class="sxs-lookup"><span data-stu-id="96af9-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="96af9-217">String</span><span class="sxs-lookup"><span data-stu-id="96af9-217">String</span></span>|<span data-ttu-id="96af9-218">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="96af9-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="96af9-219">String</span><span class="sxs-lookup"><span data-stu-id="96af9-219">String</span></span>|<span data-ttu-id="96af9-220">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="96af9-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96af9-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="96af9-221">Requirements</span></span>

|<span data-ttu-id="96af9-222">要求</span><span class="sxs-lookup"><span data-stu-id="96af9-222">Requirement</span></span>| <span data-ttu-id="96af9-223">值</span><span class="sxs-lookup"><span data-stu-id="96af9-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="96af9-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96af9-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96af9-225">1.1</span><span class="sxs-lookup"><span data-stu-id="96af9-225">1.1</span></span>|
|[<span data-ttu-id="96af9-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96af9-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96af9-227">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96af9-227">Compose or Read</span></span>|
