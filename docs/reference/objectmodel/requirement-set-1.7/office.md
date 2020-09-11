---
title: Office 命名空间-要求集1。7
description: 使用邮箱 API 要求集1.7 的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: ed98cada1328c32caa79279981bd0ce555a17385
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431393"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="dc347-103">Office (邮箱要求集 1.7) </span><span class="sxs-lookup"><span data-stu-id="dc347-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="dc347-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="dc347-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc347-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc347-106">Requirements</span></span>

|<span data-ttu-id="dc347-107">要求</span><span class="sxs-lookup"><span data-stu-id="dc347-107">Requirement</span></span>| <span data-ttu-id="dc347-108">值</span><span class="sxs-lookup"><span data-stu-id="dc347-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc347-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc347-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc347-110">1.1</span><span class="sxs-lookup"><span data-stu-id="dc347-110">1.1</span></span>|
|[<span data-ttu-id="dc347-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc347-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc347-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dc347-113">属性</span><span class="sxs-lookup"><span data-stu-id="dc347-113">Properties</span></span>

| <span data-ttu-id="dc347-114">属性</span><span class="sxs-lookup"><span data-stu-id="dc347-114">Property</span></span> | <span data-ttu-id="dc347-115">型号</span><span class="sxs-lookup"><span data-stu-id="dc347-115">Modes</span></span> | <span data-ttu-id="dc347-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="dc347-116">Return type</span></span> | <span data-ttu-id="dc347-117">最小值</span><span class="sxs-lookup"><span data-stu-id="dc347-117">Minimum</span></span><br><span data-ttu-id="dc347-118">要求集</span><span class="sxs-lookup"><span data-stu-id="dc347-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dc347-119">context</span><span class="sxs-lookup"><span data-stu-id="dc347-119">context</span></span>](office.context.md) | <span data-ttu-id="dc347-120">撰写</span><span class="sxs-lookup"><span data-stu-id="dc347-120">Compose</span></span><br><span data-ttu-id="dc347-121">阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-121">Read</span></span> | [<span data-ttu-id="dc347-122">Context</span><span class="sxs-lookup"><span data-stu-id="dc347-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="dc347-123">1.1</span><span class="sxs-lookup"><span data-stu-id="dc347-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="dc347-124">枚举</span><span class="sxs-lookup"><span data-stu-id="dc347-124">Enumerations</span></span>

| <span data-ttu-id="dc347-125">枚举</span><span class="sxs-lookup"><span data-stu-id="dc347-125">Enumeration</span></span> | <span data-ttu-id="dc347-126">型号</span><span class="sxs-lookup"><span data-stu-id="dc347-126">Modes</span></span> | <span data-ttu-id="dc347-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="dc347-127">Return type</span></span> | <span data-ttu-id="dc347-128">最小值</span><span class="sxs-lookup"><span data-stu-id="dc347-128">Minimum</span></span><br><span data-ttu-id="dc347-129">要求集</span><span class="sxs-lookup"><span data-stu-id="dc347-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dc347-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="dc347-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="dc347-131">撰写</span><span class="sxs-lookup"><span data-stu-id="dc347-131">Compose</span></span><br><span data-ttu-id="dc347-132">阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-132">Read</span></span> | <span data-ttu-id="dc347-133">String</span><span class="sxs-lookup"><span data-stu-id="dc347-133">String</span></span> | [<span data-ttu-id="dc347-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dc347-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dc347-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="dc347-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="dc347-136">撰写</span><span class="sxs-lookup"><span data-stu-id="dc347-136">Compose</span></span><br><span data-ttu-id="dc347-137">阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-137">Read</span></span> | <span data-ttu-id="dc347-138">String</span><span class="sxs-lookup"><span data-stu-id="dc347-138">String</span></span> | [<span data-ttu-id="dc347-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dc347-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dc347-140">EventType</span><span class="sxs-lookup"><span data-stu-id="dc347-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="dc347-141">撰写</span><span class="sxs-lookup"><span data-stu-id="dc347-141">Compose</span></span><br><span data-ttu-id="dc347-142">阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-142">Read</span></span> | <span data-ttu-id="dc347-143">String</span><span class="sxs-lookup"><span data-stu-id="dc347-143">String</span></span> | [<span data-ttu-id="dc347-144">1.5</span><span class="sxs-lookup"><span data-stu-id="dc347-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dc347-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="dc347-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="dc347-146">撰写</span><span class="sxs-lookup"><span data-stu-id="dc347-146">Compose</span></span><br><span data-ttu-id="dc347-147">阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-147">Read</span></span> | <span data-ttu-id="dc347-148">String</span><span class="sxs-lookup"><span data-stu-id="dc347-148">String</span></span> | [<span data-ttu-id="dc347-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dc347-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="dc347-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="dc347-150">Namespaces</span></span>

<span data-ttu-id="dc347-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true)：包含许多特定于 Outlook 的枚举，例如、、、、、 `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` 和 `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="dc347-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="dc347-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="dc347-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="dc347-153">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="dc347-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="dc347-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="dc347-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dc347-155">类型</span><span class="sxs-lookup"><span data-stu-id="dc347-155">Type</span></span>

*   <span data-ttu-id="dc347-156">String</span><span class="sxs-lookup"><span data-stu-id="dc347-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc347-157">属性：</span><span class="sxs-lookup"><span data-stu-id="dc347-157">Properties:</span></span>

|<span data-ttu-id="dc347-158">名称</span><span class="sxs-lookup"><span data-stu-id="dc347-158">Name</span></span>| <span data-ttu-id="dc347-159">类型</span><span class="sxs-lookup"><span data-stu-id="dc347-159">Type</span></span>| <span data-ttu-id="dc347-160">说明</span><span class="sxs-lookup"><span data-stu-id="dc347-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dc347-161">String</span><span class="sxs-lookup"><span data-stu-id="dc347-161">String</span></span>|<span data-ttu-id="dc347-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="dc347-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dc347-163">字符串</span><span class="sxs-lookup"><span data-stu-id="dc347-163">String</span></span>|<span data-ttu-id="dc347-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="dc347-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc347-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc347-165">Requirements</span></span>

|<span data-ttu-id="dc347-166">要求</span><span class="sxs-lookup"><span data-stu-id="dc347-166">Requirement</span></span>| <span data-ttu-id="dc347-167">值</span><span class="sxs-lookup"><span data-stu-id="dc347-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc347-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc347-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc347-169">1.1</span><span class="sxs-lookup"><span data-stu-id="dc347-169">1.1</span></span>|
|[<span data-ttu-id="dc347-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc347-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc347-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="dc347-172">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="dc347-172">CoercionType: String</span></span>

<span data-ttu-id="dc347-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="dc347-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc347-174">类型</span><span class="sxs-lookup"><span data-stu-id="dc347-174">Type</span></span>

*   <span data-ttu-id="dc347-175">String</span><span class="sxs-lookup"><span data-stu-id="dc347-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc347-176">属性：</span><span class="sxs-lookup"><span data-stu-id="dc347-176">Properties:</span></span>

|<span data-ttu-id="dc347-177">名称</span><span class="sxs-lookup"><span data-stu-id="dc347-177">Name</span></span>| <span data-ttu-id="dc347-178">类型</span><span class="sxs-lookup"><span data-stu-id="dc347-178">Type</span></span>| <span data-ttu-id="dc347-179">说明</span><span class="sxs-lookup"><span data-stu-id="dc347-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dc347-180">String</span><span class="sxs-lookup"><span data-stu-id="dc347-180">String</span></span>|<span data-ttu-id="dc347-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="dc347-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dc347-182">字符串</span><span class="sxs-lookup"><span data-stu-id="dc347-182">String</span></span>|<span data-ttu-id="dc347-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="dc347-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc347-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc347-184">Requirements</span></span>

|<span data-ttu-id="dc347-185">要求</span><span class="sxs-lookup"><span data-stu-id="dc347-185">Requirement</span></span>| <span data-ttu-id="dc347-186">值</span><span class="sxs-lookup"><span data-stu-id="dc347-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc347-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc347-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc347-188">1.1</span><span class="sxs-lookup"><span data-stu-id="dc347-188">1.1</span></span>|
|[<span data-ttu-id="dc347-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc347-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc347-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="dc347-191">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="dc347-191">EventType: String</span></span>

<span data-ttu-id="dc347-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="dc347-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="dc347-193">类型</span><span class="sxs-lookup"><span data-stu-id="dc347-193">Type</span></span>

*   <span data-ttu-id="dc347-194">String</span><span class="sxs-lookup"><span data-stu-id="dc347-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc347-195">属性：</span><span class="sxs-lookup"><span data-stu-id="dc347-195">Properties:</span></span>

| <span data-ttu-id="dc347-196">名称</span><span class="sxs-lookup"><span data-stu-id="dc347-196">Name</span></span> | <span data-ttu-id="dc347-197">类型</span><span class="sxs-lookup"><span data-stu-id="dc347-197">Type</span></span> | <span data-ttu-id="dc347-198">Description</span><span class="sxs-lookup"><span data-stu-id="dc347-198">Description</span></span> | <span data-ttu-id="dc347-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="dc347-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="dc347-200">字符串</span><span class="sxs-lookup"><span data-stu-id="dc347-200">String</span></span> | <span data-ttu-id="dc347-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="dc347-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="dc347-202">1.7</span><span class="sxs-lookup"><span data-stu-id="dc347-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="dc347-203">字符串</span><span class="sxs-lookup"><span data-stu-id="dc347-203">String</span></span> | <span data-ttu-id="dc347-204">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="dc347-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="dc347-205">1.5</span><span class="sxs-lookup"><span data-stu-id="dc347-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="dc347-206">字符串</span><span class="sxs-lookup"><span data-stu-id="dc347-206">String</span></span> | <span data-ttu-id="dc347-207">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="dc347-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="dc347-208">1.7</span><span class="sxs-lookup"><span data-stu-id="dc347-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="dc347-209">字符串</span><span class="sxs-lookup"><span data-stu-id="dc347-209">String</span></span> | <span data-ttu-id="dc347-210">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="dc347-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="dc347-211">1.7</span><span class="sxs-lookup"><span data-stu-id="dc347-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dc347-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc347-212">Requirements</span></span>

|<span data-ttu-id="dc347-213">要求</span><span class="sxs-lookup"><span data-stu-id="dc347-213">Requirement</span></span>| <span data-ttu-id="dc347-214">值</span><span class="sxs-lookup"><span data-stu-id="dc347-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc347-215">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc347-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc347-216">1.5</span><span class="sxs-lookup"><span data-stu-id="dc347-216">1.5</span></span> |
|[<span data-ttu-id="dc347-217">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc347-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc347-218">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="dc347-219">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="dc347-219">SourceProperty: String</span></span>

<span data-ttu-id="dc347-220">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="dc347-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc347-221">类型</span><span class="sxs-lookup"><span data-stu-id="dc347-221">Type</span></span>

*   <span data-ttu-id="dc347-222">String</span><span class="sxs-lookup"><span data-stu-id="dc347-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc347-223">属性：</span><span class="sxs-lookup"><span data-stu-id="dc347-223">Properties:</span></span>

|<span data-ttu-id="dc347-224">名称</span><span class="sxs-lookup"><span data-stu-id="dc347-224">Name</span></span>| <span data-ttu-id="dc347-225">类型</span><span class="sxs-lookup"><span data-stu-id="dc347-225">Type</span></span>| <span data-ttu-id="dc347-226">说明</span><span class="sxs-lookup"><span data-stu-id="dc347-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dc347-227">String</span><span class="sxs-lookup"><span data-stu-id="dc347-227">String</span></span>|<span data-ttu-id="dc347-228">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="dc347-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dc347-229">String</span><span class="sxs-lookup"><span data-stu-id="dc347-229">String</span></span>|<span data-ttu-id="dc347-230">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="dc347-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc347-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc347-231">Requirements</span></span>

|<span data-ttu-id="dc347-232">要求</span><span class="sxs-lookup"><span data-stu-id="dc347-232">Requirement</span></span>| <span data-ttu-id="dc347-233">值</span><span class="sxs-lookup"><span data-stu-id="dc347-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc347-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc347-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc347-235">1.1</span><span class="sxs-lookup"><span data-stu-id="dc347-235">1.1</span></span>|
|[<span data-ttu-id="dc347-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc347-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc347-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc347-237">Compose or Read</span></span>|
