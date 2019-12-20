---
title: Office 命名空间-要求集1。7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9bfff9c45cb157d2dcd42997a01f5ada40aecfa0
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814568"
---
# <a name="office"></a><span data-ttu-id="6004f-102">Office</span><span class="sxs-lookup"><span data-stu-id="6004f-102">Office</span></span>

<span data-ttu-id="6004f-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="6004f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6004f-105">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-105">Requirements</span></span>

|<span data-ttu-id="6004f-106">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-106">Requirement</span></span>| <span data-ttu-id="6004f-107">值</span><span class="sxs-lookup"><span data-stu-id="6004f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6004f-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6004f-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6004f-109">1.1</span><span class="sxs-lookup"><span data-stu-id="6004f-109">1.1</span></span>|
|[<span data-ttu-id="6004f-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6004f-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6004f-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6004f-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="6004f-112">属性</span><span class="sxs-lookup"><span data-stu-id="6004f-112">Properties</span></span>

| <span data-ttu-id="6004f-113">属性</span><span class="sxs-lookup"><span data-stu-id="6004f-113">Property</span></span> | <span data-ttu-id="6004f-114">型号</span><span class="sxs-lookup"><span data-stu-id="6004f-114">Modes</span></span> | <span data-ttu-id="6004f-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="6004f-115">Return type</span></span> | <span data-ttu-id="6004f-116">最低</span><span class="sxs-lookup"><span data-stu-id="6004f-116">Minimum</span></span><br><span data-ttu-id="6004f-117">要求集</span><span class="sxs-lookup"><span data-stu-id="6004f-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6004f-118">context</span><span class="sxs-lookup"><span data-stu-id="6004f-118">context</span></span>](office.context.md) | <span data-ttu-id="6004f-119">撰写</span><span class="sxs-lookup"><span data-stu-id="6004f-119">Compose</span></span><br><span data-ttu-id="6004f-120">读取</span><span class="sxs-lookup"><span data-stu-id="6004f-120">Read</span></span> | [<span data-ttu-id="6004f-121">Context</span><span class="sxs-lookup"><span data-stu-id="6004f-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="6004f-122">1.1</span><span class="sxs-lookup"><span data-stu-id="6004f-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="6004f-123">枚举</span><span class="sxs-lookup"><span data-stu-id="6004f-123">Enumerations</span></span>

| <span data-ttu-id="6004f-124">枚举</span><span class="sxs-lookup"><span data-stu-id="6004f-124">Enumeration</span></span> | <span data-ttu-id="6004f-125">型号</span><span class="sxs-lookup"><span data-stu-id="6004f-125">Modes</span></span> | <span data-ttu-id="6004f-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="6004f-126">Return type</span></span> | <span data-ttu-id="6004f-127">最低</span><span class="sxs-lookup"><span data-stu-id="6004f-127">Minimum</span></span><br><span data-ttu-id="6004f-128">要求集</span><span class="sxs-lookup"><span data-stu-id="6004f-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6004f-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6004f-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6004f-130">撰写</span><span class="sxs-lookup"><span data-stu-id="6004f-130">Compose</span></span><br><span data-ttu-id="6004f-131">读取</span><span class="sxs-lookup"><span data-stu-id="6004f-131">Read</span></span> | <span data-ttu-id="6004f-132">String</span><span class="sxs-lookup"><span data-stu-id="6004f-132">String</span></span> | [<span data-ttu-id="6004f-133">1.1</span><span class="sxs-lookup"><span data-stu-id="6004f-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6004f-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6004f-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6004f-135">撰写</span><span class="sxs-lookup"><span data-stu-id="6004f-135">Compose</span></span><br><span data-ttu-id="6004f-136">读取</span><span class="sxs-lookup"><span data-stu-id="6004f-136">Read</span></span> | <span data-ttu-id="6004f-137">String</span><span class="sxs-lookup"><span data-stu-id="6004f-137">String</span></span> | [<span data-ttu-id="6004f-138">1.1</span><span class="sxs-lookup"><span data-stu-id="6004f-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6004f-139">EventType</span><span class="sxs-lookup"><span data-stu-id="6004f-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6004f-140">撰写</span><span class="sxs-lookup"><span data-stu-id="6004f-140">Compose</span></span><br><span data-ttu-id="6004f-141">读取</span><span class="sxs-lookup"><span data-stu-id="6004f-141">Read</span></span> | <span data-ttu-id="6004f-142">String</span><span class="sxs-lookup"><span data-stu-id="6004f-142">String</span></span> | [<span data-ttu-id="6004f-143">1.5</span><span class="sxs-lookup"><span data-stu-id="6004f-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6004f-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6004f-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6004f-145">撰写</span><span class="sxs-lookup"><span data-stu-id="6004f-145">Compose</span></span><br><span data-ttu-id="6004f-146">读取</span><span class="sxs-lookup"><span data-stu-id="6004f-146">Read</span></span> | <span data-ttu-id="6004f-147">String</span><span class="sxs-lookup"><span data-stu-id="6004f-147">String</span></span> | [<span data-ttu-id="6004f-148">1.1</span><span class="sxs-lookup"><span data-stu-id="6004f-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="6004f-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="6004f-149">Namespaces</span></span>

<span data-ttu-id="6004f-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="6004f-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="6004f-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="6004f-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="6004f-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="6004f-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="6004f-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="6004f-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6004f-154">类型</span><span class="sxs-lookup"><span data-stu-id="6004f-154">Type</span></span>

*   <span data-ttu-id="6004f-155">String</span><span class="sxs-lookup"><span data-stu-id="6004f-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6004f-156">属性：</span><span class="sxs-lookup"><span data-stu-id="6004f-156">Properties:</span></span>

|<span data-ttu-id="6004f-157">名称</span><span class="sxs-lookup"><span data-stu-id="6004f-157">Name</span></span>| <span data-ttu-id="6004f-158">类型</span><span class="sxs-lookup"><span data-stu-id="6004f-158">Type</span></span>| <span data-ttu-id="6004f-159">说明</span><span class="sxs-lookup"><span data-stu-id="6004f-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6004f-160">String</span><span class="sxs-lookup"><span data-stu-id="6004f-160">String</span></span>|<span data-ttu-id="6004f-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="6004f-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6004f-162">String</span><span class="sxs-lookup"><span data-stu-id="6004f-162">String</span></span>|<span data-ttu-id="6004f-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="6004f-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6004f-164">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-164">Requirements</span></span>

|<span data-ttu-id="6004f-165">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-165">Requirement</span></span>| <span data-ttu-id="6004f-166">值</span><span class="sxs-lookup"><span data-stu-id="6004f-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="6004f-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6004f-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6004f-168">1.1</span><span class="sxs-lookup"><span data-stu-id="6004f-168">1.1</span></span>|
|[<span data-ttu-id="6004f-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6004f-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6004f-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6004f-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="6004f-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="6004f-171">CoercionType: String</span></span>

<span data-ttu-id="6004f-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="6004f-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6004f-173">类型</span><span class="sxs-lookup"><span data-stu-id="6004f-173">Type</span></span>

*   <span data-ttu-id="6004f-174">String</span><span class="sxs-lookup"><span data-stu-id="6004f-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6004f-175">属性：</span><span class="sxs-lookup"><span data-stu-id="6004f-175">Properties:</span></span>

|<span data-ttu-id="6004f-176">名称</span><span class="sxs-lookup"><span data-stu-id="6004f-176">Name</span></span>| <span data-ttu-id="6004f-177">类型</span><span class="sxs-lookup"><span data-stu-id="6004f-177">Type</span></span>| <span data-ttu-id="6004f-178">说明</span><span class="sxs-lookup"><span data-stu-id="6004f-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6004f-179">String</span><span class="sxs-lookup"><span data-stu-id="6004f-179">String</span></span>|<span data-ttu-id="6004f-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="6004f-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6004f-181">String</span><span class="sxs-lookup"><span data-stu-id="6004f-181">String</span></span>|<span data-ttu-id="6004f-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="6004f-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6004f-183">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-183">Requirements</span></span>

|<span data-ttu-id="6004f-184">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-184">Requirement</span></span>| <span data-ttu-id="6004f-185">值</span><span class="sxs-lookup"><span data-stu-id="6004f-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="6004f-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6004f-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6004f-187">1.1</span><span class="sxs-lookup"><span data-stu-id="6004f-187">1.1</span></span>|
|[<span data-ttu-id="6004f-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6004f-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6004f-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6004f-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="6004f-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="6004f-190">EventType: String</span></span>

<span data-ttu-id="6004f-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="6004f-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6004f-192">类型</span><span class="sxs-lookup"><span data-stu-id="6004f-192">Type</span></span>

*   <span data-ttu-id="6004f-193">String</span><span class="sxs-lookup"><span data-stu-id="6004f-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6004f-194">属性：</span><span class="sxs-lookup"><span data-stu-id="6004f-194">Properties:</span></span>

| <span data-ttu-id="6004f-195">名称</span><span class="sxs-lookup"><span data-stu-id="6004f-195">Name</span></span> | <span data-ttu-id="6004f-196">类型</span><span class="sxs-lookup"><span data-stu-id="6004f-196">Type</span></span> | <span data-ttu-id="6004f-197">说明</span><span class="sxs-lookup"><span data-stu-id="6004f-197">Description</span></span> | <span data-ttu-id="6004f-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="6004f-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="6004f-199">String</span><span class="sxs-lookup"><span data-stu-id="6004f-199">String</span></span> | <span data-ttu-id="6004f-200">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="6004f-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="6004f-201">1.7</span><span class="sxs-lookup"><span data-stu-id="6004f-201">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="6004f-202">String</span><span class="sxs-lookup"><span data-stu-id="6004f-202">String</span></span> | <span data-ttu-id="6004f-203">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="6004f-203">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="6004f-204">1.5</span><span class="sxs-lookup"><span data-stu-id="6004f-204">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="6004f-205">String</span><span class="sxs-lookup"><span data-stu-id="6004f-205">String</span></span> | <span data-ttu-id="6004f-206">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="6004f-206">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="6004f-207">1.7</span><span class="sxs-lookup"><span data-stu-id="6004f-207">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="6004f-208">String</span><span class="sxs-lookup"><span data-stu-id="6004f-208">String</span></span> | <span data-ttu-id="6004f-209">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="6004f-209">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="6004f-210">1.7</span><span class="sxs-lookup"><span data-stu-id="6004f-210">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6004f-211">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-211">Requirements</span></span>

|<span data-ttu-id="6004f-212">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-212">Requirement</span></span>| <span data-ttu-id="6004f-213">值</span><span class="sxs-lookup"><span data-stu-id="6004f-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="6004f-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6004f-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6004f-215">1.5</span><span class="sxs-lookup"><span data-stu-id="6004f-215">1.5</span></span> |
|[<span data-ttu-id="6004f-216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6004f-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6004f-217">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6004f-217">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="6004f-218">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="6004f-218">SourceProperty: String</span></span>

<span data-ttu-id="6004f-219">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="6004f-219">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6004f-220">类型</span><span class="sxs-lookup"><span data-stu-id="6004f-220">Type</span></span>

*   <span data-ttu-id="6004f-221">String</span><span class="sxs-lookup"><span data-stu-id="6004f-221">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6004f-222">属性：</span><span class="sxs-lookup"><span data-stu-id="6004f-222">Properties:</span></span>

|<span data-ttu-id="6004f-223">名称</span><span class="sxs-lookup"><span data-stu-id="6004f-223">Name</span></span>| <span data-ttu-id="6004f-224">类型</span><span class="sxs-lookup"><span data-stu-id="6004f-224">Type</span></span>| <span data-ttu-id="6004f-225">说明</span><span class="sxs-lookup"><span data-stu-id="6004f-225">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6004f-226">String</span><span class="sxs-lookup"><span data-stu-id="6004f-226">String</span></span>|<span data-ttu-id="6004f-227">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="6004f-227">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6004f-228">String</span><span class="sxs-lookup"><span data-stu-id="6004f-228">String</span></span>|<span data-ttu-id="6004f-229">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="6004f-229">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6004f-230">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-230">Requirements</span></span>

|<span data-ttu-id="6004f-231">要求</span><span class="sxs-lookup"><span data-stu-id="6004f-231">Requirement</span></span>| <span data-ttu-id="6004f-232">值</span><span class="sxs-lookup"><span data-stu-id="6004f-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="6004f-233">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6004f-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6004f-234">1.1</span><span class="sxs-lookup"><span data-stu-id="6004f-234">1.1</span></span>|
|[<span data-ttu-id="6004f-235">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6004f-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6004f-236">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6004f-236">Compose or Read</span></span>|
