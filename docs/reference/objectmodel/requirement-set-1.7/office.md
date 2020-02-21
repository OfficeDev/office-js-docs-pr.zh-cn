---
title: Office 命名空间-要求集1。7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 23f3fb705c03eabd8ee7fce53f4c89a48128672f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165346"
---
# <a name="office"></a><span data-ttu-id="95adb-102">Office</span><span class="sxs-lookup"><span data-stu-id="95adb-102">Office</span></span>

<span data-ttu-id="95adb-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="95adb-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="95adb-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="95adb-105">Requirements</span></span>

|<span data-ttu-id="95adb-106">要求</span><span class="sxs-lookup"><span data-stu-id="95adb-106">Requirement</span></span>| <span data-ttu-id="95adb-107">值</span><span class="sxs-lookup"><span data-stu-id="95adb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="95adb-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95adb-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95adb-109">1.1</span><span class="sxs-lookup"><span data-stu-id="95adb-109">1.1</span></span>|
|[<span data-ttu-id="95adb-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95adb-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95adb-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95adb-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="95adb-112">属性</span><span class="sxs-lookup"><span data-stu-id="95adb-112">Properties</span></span>

| <span data-ttu-id="95adb-113">属性</span><span class="sxs-lookup"><span data-stu-id="95adb-113">Property</span></span> | <span data-ttu-id="95adb-114">型号</span><span class="sxs-lookup"><span data-stu-id="95adb-114">Modes</span></span> | <span data-ttu-id="95adb-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="95adb-115">Return type</span></span> | <span data-ttu-id="95adb-116">最低</span><span class="sxs-lookup"><span data-stu-id="95adb-116">Minimum</span></span><br><span data-ttu-id="95adb-117">要求集</span><span class="sxs-lookup"><span data-stu-id="95adb-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="95adb-118">context</span><span class="sxs-lookup"><span data-stu-id="95adb-118">context</span></span>](office.context.md) | <span data-ttu-id="95adb-119">撰写</span><span class="sxs-lookup"><span data-stu-id="95adb-119">Compose</span></span><br><span data-ttu-id="95adb-120">读取</span><span class="sxs-lookup"><span data-stu-id="95adb-120">Read</span></span> | [<span data-ttu-id="95adb-121">Context</span><span class="sxs-lookup"><span data-stu-id="95adb-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="95adb-122">1.1</span><span class="sxs-lookup"><span data-stu-id="95adb-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="95adb-123">枚举</span><span class="sxs-lookup"><span data-stu-id="95adb-123">Enumerations</span></span>

| <span data-ttu-id="95adb-124">枚举</span><span class="sxs-lookup"><span data-stu-id="95adb-124">Enumeration</span></span> | <span data-ttu-id="95adb-125">型号</span><span class="sxs-lookup"><span data-stu-id="95adb-125">Modes</span></span> | <span data-ttu-id="95adb-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="95adb-126">Return type</span></span> | <span data-ttu-id="95adb-127">最低</span><span class="sxs-lookup"><span data-stu-id="95adb-127">Minimum</span></span><br><span data-ttu-id="95adb-128">要求集</span><span class="sxs-lookup"><span data-stu-id="95adb-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="95adb-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="95adb-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="95adb-130">撰写</span><span class="sxs-lookup"><span data-stu-id="95adb-130">Compose</span></span><br><span data-ttu-id="95adb-131">读取</span><span class="sxs-lookup"><span data-stu-id="95adb-131">Read</span></span> | <span data-ttu-id="95adb-132">String</span><span class="sxs-lookup"><span data-stu-id="95adb-132">String</span></span> | [<span data-ttu-id="95adb-133">1.1</span><span class="sxs-lookup"><span data-stu-id="95adb-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95adb-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="95adb-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="95adb-135">撰写</span><span class="sxs-lookup"><span data-stu-id="95adb-135">Compose</span></span><br><span data-ttu-id="95adb-136">读取</span><span class="sxs-lookup"><span data-stu-id="95adb-136">Read</span></span> | <span data-ttu-id="95adb-137">String</span><span class="sxs-lookup"><span data-stu-id="95adb-137">String</span></span> | [<span data-ttu-id="95adb-138">1.1</span><span class="sxs-lookup"><span data-stu-id="95adb-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95adb-139">EventType</span><span class="sxs-lookup"><span data-stu-id="95adb-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="95adb-140">撰写</span><span class="sxs-lookup"><span data-stu-id="95adb-140">Compose</span></span><br><span data-ttu-id="95adb-141">读取</span><span class="sxs-lookup"><span data-stu-id="95adb-141">Read</span></span> | <span data-ttu-id="95adb-142">String</span><span class="sxs-lookup"><span data-stu-id="95adb-142">String</span></span> | [<span data-ttu-id="95adb-143">1.5</span><span class="sxs-lookup"><span data-stu-id="95adb-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="95adb-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="95adb-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="95adb-145">撰写</span><span class="sxs-lookup"><span data-stu-id="95adb-145">Compose</span></span><br><span data-ttu-id="95adb-146">读取</span><span class="sxs-lookup"><span data-stu-id="95adb-146">Read</span></span> | <span data-ttu-id="95adb-147">String</span><span class="sxs-lookup"><span data-stu-id="95adb-147">String</span></span> | [<span data-ttu-id="95adb-148">1.1</span><span class="sxs-lookup"><span data-stu-id="95adb-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="95adb-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="95adb-149">Namespaces</span></span>

<span data-ttu-id="95adb-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="95adb-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="95adb-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="95adb-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="95adb-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="95adb-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="95adb-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="95adb-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="95adb-154">类型</span><span class="sxs-lookup"><span data-stu-id="95adb-154">Type</span></span>

*   <span data-ttu-id="95adb-155">String</span><span class="sxs-lookup"><span data-stu-id="95adb-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="95adb-156">属性：</span><span class="sxs-lookup"><span data-stu-id="95adb-156">Properties:</span></span>

|<span data-ttu-id="95adb-157">名称</span><span class="sxs-lookup"><span data-stu-id="95adb-157">Name</span></span>| <span data-ttu-id="95adb-158">类型</span><span class="sxs-lookup"><span data-stu-id="95adb-158">Type</span></span>| <span data-ttu-id="95adb-159">说明</span><span class="sxs-lookup"><span data-stu-id="95adb-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="95adb-160">String</span><span class="sxs-lookup"><span data-stu-id="95adb-160">String</span></span>|<span data-ttu-id="95adb-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="95adb-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="95adb-162">String</span><span class="sxs-lookup"><span data-stu-id="95adb-162">String</span></span>|<span data-ttu-id="95adb-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="95adb-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95adb-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="95adb-164">Requirements</span></span>

|<span data-ttu-id="95adb-165">要求</span><span class="sxs-lookup"><span data-stu-id="95adb-165">Requirement</span></span>| <span data-ttu-id="95adb-166">值</span><span class="sxs-lookup"><span data-stu-id="95adb-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="95adb-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95adb-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95adb-168">1.1</span><span class="sxs-lookup"><span data-stu-id="95adb-168">1.1</span></span>|
|[<span data-ttu-id="95adb-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95adb-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95adb-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95adb-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="95adb-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="95adb-171">CoercionType: String</span></span>

<span data-ttu-id="95adb-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="95adb-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="95adb-173">类型</span><span class="sxs-lookup"><span data-stu-id="95adb-173">Type</span></span>

*   <span data-ttu-id="95adb-174">String</span><span class="sxs-lookup"><span data-stu-id="95adb-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="95adb-175">属性：</span><span class="sxs-lookup"><span data-stu-id="95adb-175">Properties:</span></span>

|<span data-ttu-id="95adb-176">名称</span><span class="sxs-lookup"><span data-stu-id="95adb-176">Name</span></span>| <span data-ttu-id="95adb-177">类型</span><span class="sxs-lookup"><span data-stu-id="95adb-177">Type</span></span>| <span data-ttu-id="95adb-178">说明</span><span class="sxs-lookup"><span data-stu-id="95adb-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="95adb-179">String</span><span class="sxs-lookup"><span data-stu-id="95adb-179">String</span></span>|<span data-ttu-id="95adb-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="95adb-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="95adb-181">String</span><span class="sxs-lookup"><span data-stu-id="95adb-181">String</span></span>|<span data-ttu-id="95adb-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="95adb-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95adb-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="95adb-183">Requirements</span></span>

|<span data-ttu-id="95adb-184">要求</span><span class="sxs-lookup"><span data-stu-id="95adb-184">Requirement</span></span>| <span data-ttu-id="95adb-185">值</span><span class="sxs-lookup"><span data-stu-id="95adb-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="95adb-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95adb-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95adb-187">1.1</span><span class="sxs-lookup"><span data-stu-id="95adb-187">1.1</span></span>|
|[<span data-ttu-id="95adb-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95adb-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95adb-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95adb-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="95adb-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="95adb-190">EventType: String</span></span>

<span data-ttu-id="95adb-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="95adb-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="95adb-192">类型</span><span class="sxs-lookup"><span data-stu-id="95adb-192">Type</span></span>

*   <span data-ttu-id="95adb-193">String</span><span class="sxs-lookup"><span data-stu-id="95adb-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="95adb-194">属性：</span><span class="sxs-lookup"><span data-stu-id="95adb-194">Properties:</span></span>

| <span data-ttu-id="95adb-195">名称</span><span class="sxs-lookup"><span data-stu-id="95adb-195">Name</span></span> | <span data-ttu-id="95adb-196">类型</span><span class="sxs-lookup"><span data-stu-id="95adb-196">Type</span></span> | <span data-ttu-id="95adb-197">说明</span><span class="sxs-lookup"><span data-stu-id="95adb-197">Description</span></span> | <span data-ttu-id="95adb-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="95adb-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="95adb-199">String</span><span class="sxs-lookup"><span data-stu-id="95adb-199">String</span></span> | <span data-ttu-id="95adb-200">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="95adb-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="95adb-201">1.7</span><span class="sxs-lookup"><span data-stu-id="95adb-201">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="95adb-202">String</span><span class="sxs-lookup"><span data-stu-id="95adb-202">String</span></span> | <span data-ttu-id="95adb-203">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="95adb-203">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="95adb-204">1.5</span><span class="sxs-lookup"><span data-stu-id="95adb-204">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="95adb-205">String</span><span class="sxs-lookup"><span data-stu-id="95adb-205">String</span></span> | <span data-ttu-id="95adb-206">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="95adb-206">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="95adb-207">1.7</span><span class="sxs-lookup"><span data-stu-id="95adb-207">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="95adb-208">String</span><span class="sxs-lookup"><span data-stu-id="95adb-208">String</span></span> | <span data-ttu-id="95adb-209">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="95adb-209">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="95adb-210">1.7</span><span class="sxs-lookup"><span data-stu-id="95adb-210">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="95adb-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="95adb-211">Requirements</span></span>

|<span data-ttu-id="95adb-212">要求</span><span class="sxs-lookup"><span data-stu-id="95adb-212">Requirement</span></span>| <span data-ttu-id="95adb-213">值</span><span class="sxs-lookup"><span data-stu-id="95adb-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="95adb-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95adb-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95adb-215">1.5</span><span class="sxs-lookup"><span data-stu-id="95adb-215">1.5</span></span> |
|[<span data-ttu-id="95adb-216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95adb-216">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95adb-217">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95adb-217">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="95adb-218">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="95adb-218">SourceProperty: String</span></span>

<span data-ttu-id="95adb-219">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="95adb-219">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="95adb-220">类型</span><span class="sxs-lookup"><span data-stu-id="95adb-220">Type</span></span>

*   <span data-ttu-id="95adb-221">String</span><span class="sxs-lookup"><span data-stu-id="95adb-221">String</span></span>

##### <a name="properties"></a><span data-ttu-id="95adb-222">属性：</span><span class="sxs-lookup"><span data-stu-id="95adb-222">Properties:</span></span>

|<span data-ttu-id="95adb-223">名称</span><span class="sxs-lookup"><span data-stu-id="95adb-223">Name</span></span>| <span data-ttu-id="95adb-224">类型</span><span class="sxs-lookup"><span data-stu-id="95adb-224">Type</span></span>| <span data-ttu-id="95adb-225">说明</span><span class="sxs-lookup"><span data-stu-id="95adb-225">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="95adb-226">String</span><span class="sxs-lookup"><span data-stu-id="95adb-226">String</span></span>|<span data-ttu-id="95adb-227">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="95adb-227">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="95adb-228">String</span><span class="sxs-lookup"><span data-stu-id="95adb-228">String</span></span>|<span data-ttu-id="95adb-229">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="95adb-229">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95adb-230">Requirements</span><span class="sxs-lookup"><span data-stu-id="95adb-230">Requirements</span></span>

|<span data-ttu-id="95adb-231">要求</span><span class="sxs-lookup"><span data-stu-id="95adb-231">Requirement</span></span>| <span data-ttu-id="95adb-232">值</span><span class="sxs-lookup"><span data-stu-id="95adb-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="95adb-233">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="95adb-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95adb-234">1.1</span><span class="sxs-lookup"><span data-stu-id="95adb-234">1.1</span></span>|
|[<span data-ttu-id="95adb-235">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="95adb-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95adb-236">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="95adb-236">Compose or Read</span></span>|
