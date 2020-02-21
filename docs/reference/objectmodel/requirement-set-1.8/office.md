---
title: Office 命名空间-要求集1。8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: c5c431f7a958f1c2a956f36e90ad0f3a205c6669
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163624"
---
# <a name="office"></a><span data-ttu-id="35cff-102">Office</span><span class="sxs-lookup"><span data-stu-id="35cff-102">Office</span></span>

<span data-ttu-id="35cff-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="35cff-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="35cff-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="35cff-105">Requirements</span></span>

|<span data-ttu-id="35cff-106">要求</span><span class="sxs-lookup"><span data-stu-id="35cff-106">Requirement</span></span>| <span data-ttu-id="35cff-107">值</span><span class="sxs-lookup"><span data-stu-id="35cff-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="35cff-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="35cff-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35cff-109">1.1</span><span class="sxs-lookup"><span data-stu-id="35cff-109">1.1</span></span>|
|[<span data-ttu-id="35cff-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35cff-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="35cff-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="35cff-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="35cff-112">属性</span><span class="sxs-lookup"><span data-stu-id="35cff-112">Properties</span></span>

| <span data-ttu-id="35cff-113">属性</span><span class="sxs-lookup"><span data-stu-id="35cff-113">Property</span></span> | <span data-ttu-id="35cff-114">型号</span><span class="sxs-lookup"><span data-stu-id="35cff-114">Modes</span></span> | <span data-ttu-id="35cff-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="35cff-115">Return type</span></span> | <span data-ttu-id="35cff-116">最低</span><span class="sxs-lookup"><span data-stu-id="35cff-116">Minimum</span></span><br><span data-ttu-id="35cff-117">要求集</span><span class="sxs-lookup"><span data-stu-id="35cff-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="35cff-118">context</span><span class="sxs-lookup"><span data-stu-id="35cff-118">context</span></span>](office.context.md) | <span data-ttu-id="35cff-119">撰写</span><span class="sxs-lookup"><span data-stu-id="35cff-119">Compose</span></span><br><span data-ttu-id="35cff-120">读取</span><span class="sxs-lookup"><span data-stu-id="35cff-120">Read</span></span> | [<span data-ttu-id="35cff-121">Context</span><span class="sxs-lookup"><span data-stu-id="35cff-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="35cff-122">1.1</span><span class="sxs-lookup"><span data-stu-id="35cff-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="35cff-123">枚举</span><span class="sxs-lookup"><span data-stu-id="35cff-123">Enumerations</span></span>

| <span data-ttu-id="35cff-124">枚举</span><span class="sxs-lookup"><span data-stu-id="35cff-124">Enumeration</span></span> | <span data-ttu-id="35cff-125">型号</span><span class="sxs-lookup"><span data-stu-id="35cff-125">Modes</span></span> | <span data-ttu-id="35cff-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="35cff-126">Return type</span></span> | <span data-ttu-id="35cff-127">最低</span><span class="sxs-lookup"><span data-stu-id="35cff-127">Minimum</span></span><br><span data-ttu-id="35cff-128">要求集</span><span class="sxs-lookup"><span data-stu-id="35cff-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="35cff-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="35cff-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="35cff-130">撰写</span><span class="sxs-lookup"><span data-stu-id="35cff-130">Compose</span></span><br><span data-ttu-id="35cff-131">读取</span><span class="sxs-lookup"><span data-stu-id="35cff-131">Read</span></span> | <span data-ttu-id="35cff-132">String</span><span class="sxs-lookup"><span data-stu-id="35cff-132">String</span></span> | [<span data-ttu-id="35cff-133">1.1</span><span class="sxs-lookup"><span data-stu-id="35cff-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="35cff-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="35cff-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="35cff-135">撰写</span><span class="sxs-lookup"><span data-stu-id="35cff-135">Compose</span></span><br><span data-ttu-id="35cff-136">读取</span><span class="sxs-lookup"><span data-stu-id="35cff-136">Read</span></span> | <span data-ttu-id="35cff-137">String</span><span class="sxs-lookup"><span data-stu-id="35cff-137">String</span></span> | [<span data-ttu-id="35cff-138">1.1</span><span class="sxs-lookup"><span data-stu-id="35cff-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="35cff-139">EventType</span><span class="sxs-lookup"><span data-stu-id="35cff-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="35cff-140">撰写</span><span class="sxs-lookup"><span data-stu-id="35cff-140">Compose</span></span><br><span data-ttu-id="35cff-141">读取</span><span class="sxs-lookup"><span data-stu-id="35cff-141">Read</span></span> | <span data-ttu-id="35cff-142">String</span><span class="sxs-lookup"><span data-stu-id="35cff-142">String</span></span> | [<span data-ttu-id="35cff-143">1.5</span><span class="sxs-lookup"><span data-stu-id="35cff-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="35cff-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="35cff-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="35cff-145">撰写</span><span class="sxs-lookup"><span data-stu-id="35cff-145">Compose</span></span><br><span data-ttu-id="35cff-146">读取</span><span class="sxs-lookup"><span data-stu-id="35cff-146">Read</span></span> | <span data-ttu-id="35cff-147">String</span><span class="sxs-lookup"><span data-stu-id="35cff-147">String</span></span> | [<span data-ttu-id="35cff-148">1.1</span><span class="sxs-lookup"><span data-stu-id="35cff-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="35cff-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="35cff-149">Namespaces</span></span>

<span data-ttu-id="35cff-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="35cff-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="35cff-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="35cff-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="35cff-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="35cff-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="35cff-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="35cff-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="35cff-154">类型</span><span class="sxs-lookup"><span data-stu-id="35cff-154">Type</span></span>

*   <span data-ttu-id="35cff-155">String</span><span class="sxs-lookup"><span data-stu-id="35cff-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35cff-156">属性：</span><span class="sxs-lookup"><span data-stu-id="35cff-156">Properties:</span></span>

|<span data-ttu-id="35cff-157">名称</span><span class="sxs-lookup"><span data-stu-id="35cff-157">Name</span></span>| <span data-ttu-id="35cff-158">类型</span><span class="sxs-lookup"><span data-stu-id="35cff-158">Type</span></span>| <span data-ttu-id="35cff-159">说明</span><span class="sxs-lookup"><span data-stu-id="35cff-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="35cff-160">String</span><span class="sxs-lookup"><span data-stu-id="35cff-160">String</span></span>|<span data-ttu-id="35cff-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="35cff-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="35cff-162">String</span><span class="sxs-lookup"><span data-stu-id="35cff-162">String</span></span>|<span data-ttu-id="35cff-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="35cff-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35cff-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="35cff-164">Requirements</span></span>

|<span data-ttu-id="35cff-165">要求</span><span class="sxs-lookup"><span data-stu-id="35cff-165">Requirement</span></span>| <span data-ttu-id="35cff-166">值</span><span class="sxs-lookup"><span data-stu-id="35cff-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="35cff-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="35cff-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35cff-168">1.1</span><span class="sxs-lookup"><span data-stu-id="35cff-168">1.1</span></span>|
|[<span data-ttu-id="35cff-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35cff-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="35cff-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="35cff-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="35cff-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="35cff-171">CoercionType: String</span></span>

<span data-ttu-id="35cff-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="35cff-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="35cff-173">类型</span><span class="sxs-lookup"><span data-stu-id="35cff-173">Type</span></span>

*   <span data-ttu-id="35cff-174">String</span><span class="sxs-lookup"><span data-stu-id="35cff-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35cff-175">属性：</span><span class="sxs-lookup"><span data-stu-id="35cff-175">Properties:</span></span>

|<span data-ttu-id="35cff-176">名称</span><span class="sxs-lookup"><span data-stu-id="35cff-176">Name</span></span>| <span data-ttu-id="35cff-177">类型</span><span class="sxs-lookup"><span data-stu-id="35cff-177">Type</span></span>| <span data-ttu-id="35cff-178">说明</span><span class="sxs-lookup"><span data-stu-id="35cff-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="35cff-179">String</span><span class="sxs-lookup"><span data-stu-id="35cff-179">String</span></span>|<span data-ttu-id="35cff-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="35cff-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="35cff-181">String</span><span class="sxs-lookup"><span data-stu-id="35cff-181">String</span></span>|<span data-ttu-id="35cff-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="35cff-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35cff-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="35cff-183">Requirements</span></span>

|<span data-ttu-id="35cff-184">要求</span><span class="sxs-lookup"><span data-stu-id="35cff-184">Requirement</span></span>| <span data-ttu-id="35cff-185">值</span><span class="sxs-lookup"><span data-stu-id="35cff-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="35cff-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="35cff-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35cff-187">1.1</span><span class="sxs-lookup"><span data-stu-id="35cff-187">1.1</span></span>|
|[<span data-ttu-id="35cff-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35cff-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="35cff-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="35cff-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="35cff-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="35cff-190">EventType: String</span></span>

<span data-ttu-id="35cff-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="35cff-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="35cff-192">类型</span><span class="sxs-lookup"><span data-stu-id="35cff-192">Type</span></span>

*   <span data-ttu-id="35cff-193">String</span><span class="sxs-lookup"><span data-stu-id="35cff-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35cff-194">属性：</span><span class="sxs-lookup"><span data-stu-id="35cff-194">Properties:</span></span>

| <span data-ttu-id="35cff-195">名称</span><span class="sxs-lookup"><span data-stu-id="35cff-195">Name</span></span> | <span data-ttu-id="35cff-196">类型</span><span class="sxs-lookup"><span data-stu-id="35cff-196">Type</span></span> | <span data-ttu-id="35cff-197">说明</span><span class="sxs-lookup"><span data-stu-id="35cff-197">Description</span></span> | <span data-ttu-id="35cff-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="35cff-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="35cff-199">String</span><span class="sxs-lookup"><span data-stu-id="35cff-199">String</span></span> | <span data-ttu-id="35cff-200">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="35cff-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="35cff-201">1.7</span><span class="sxs-lookup"><span data-stu-id="35cff-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="35cff-202">String</span><span class="sxs-lookup"><span data-stu-id="35cff-202">String</span></span> | <span data-ttu-id="35cff-203">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="35cff-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="35cff-204">1.8</span><span class="sxs-lookup"><span data-stu-id="35cff-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="35cff-205">String</span><span class="sxs-lookup"><span data-stu-id="35cff-205">String</span></span> | <span data-ttu-id="35cff-206">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="35cff-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="35cff-207">1.8</span><span class="sxs-lookup"><span data-stu-id="35cff-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="35cff-208">String</span><span class="sxs-lookup"><span data-stu-id="35cff-208">String</span></span> | <span data-ttu-id="35cff-209">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="35cff-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="35cff-210">1.5</span><span class="sxs-lookup"><span data-stu-id="35cff-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="35cff-211">String</span><span class="sxs-lookup"><span data-stu-id="35cff-211">String</span></span> | <span data-ttu-id="35cff-212">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="35cff-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="35cff-213">1.7</span><span class="sxs-lookup"><span data-stu-id="35cff-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="35cff-214">String</span><span class="sxs-lookup"><span data-stu-id="35cff-214">String</span></span> | <span data-ttu-id="35cff-215">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="35cff-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="35cff-216">1.7</span><span class="sxs-lookup"><span data-stu-id="35cff-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35cff-217">Requirements</span><span class="sxs-lookup"><span data-stu-id="35cff-217">Requirements</span></span>

|<span data-ttu-id="35cff-218">要求</span><span class="sxs-lookup"><span data-stu-id="35cff-218">Requirement</span></span>| <span data-ttu-id="35cff-219">值</span><span class="sxs-lookup"><span data-stu-id="35cff-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="35cff-220">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="35cff-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35cff-221">1.5</span><span class="sxs-lookup"><span data-stu-id="35cff-221">1.5</span></span> |
|[<span data-ttu-id="35cff-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35cff-222">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="35cff-223">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="35cff-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="35cff-224">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="35cff-224">SourceProperty: String</span></span>

<span data-ttu-id="35cff-225">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="35cff-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="35cff-226">类型</span><span class="sxs-lookup"><span data-stu-id="35cff-226">Type</span></span>

*   <span data-ttu-id="35cff-227">String</span><span class="sxs-lookup"><span data-stu-id="35cff-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35cff-228">属性：</span><span class="sxs-lookup"><span data-stu-id="35cff-228">Properties:</span></span>

|<span data-ttu-id="35cff-229">名称</span><span class="sxs-lookup"><span data-stu-id="35cff-229">Name</span></span>| <span data-ttu-id="35cff-230">类型</span><span class="sxs-lookup"><span data-stu-id="35cff-230">Type</span></span>| <span data-ttu-id="35cff-231">说明</span><span class="sxs-lookup"><span data-stu-id="35cff-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="35cff-232">String</span><span class="sxs-lookup"><span data-stu-id="35cff-232">String</span></span>|<span data-ttu-id="35cff-233">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="35cff-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="35cff-234">String</span><span class="sxs-lookup"><span data-stu-id="35cff-234">String</span></span>|<span data-ttu-id="35cff-235">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="35cff-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35cff-236">Requirements</span><span class="sxs-lookup"><span data-stu-id="35cff-236">Requirements</span></span>

|<span data-ttu-id="35cff-237">要求</span><span class="sxs-lookup"><span data-stu-id="35cff-237">Requirement</span></span>| <span data-ttu-id="35cff-238">值</span><span class="sxs-lookup"><span data-stu-id="35cff-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="35cff-239">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="35cff-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="35cff-240">1.1</span><span class="sxs-lookup"><span data-stu-id="35cff-240">1.1</span></span>|
|[<span data-ttu-id="35cff-241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35cff-241">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="35cff-242">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="35cff-242">Compose or Read</span></span>|
