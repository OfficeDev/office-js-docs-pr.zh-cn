---
title: Office 命名空间 - 预览要求集
description: 使用邮箱 API preview 要求集的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: d72e5c78a7fd8d3c00b8f84e7d9b05ee6defc0c5
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890856"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="f70e5-103">Office （邮箱预览要求集）</span><span class="sxs-lookup"><span data-stu-id="f70e5-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="f70e5-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="f70e5-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f70e5-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="f70e5-106">Requirements</span></span>

|<span data-ttu-id="f70e5-107">要求</span><span class="sxs-lookup"><span data-stu-id="f70e5-107">Requirement</span></span>| <span data-ttu-id="f70e5-108">值</span><span class="sxs-lookup"><span data-stu-id="f70e5-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70e5-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f70e5-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f70e5-110">1.1</span><span class="sxs-lookup"><span data-stu-id="f70e5-110">1.1</span></span>|
|[<span data-ttu-id="f70e5-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f70e5-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f70e5-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f70e5-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f70e5-113">属性</span><span class="sxs-lookup"><span data-stu-id="f70e5-113">Properties</span></span>

| <span data-ttu-id="f70e5-114">属性</span><span class="sxs-lookup"><span data-stu-id="f70e5-114">Property</span></span> | <span data-ttu-id="f70e5-115">型号</span><span class="sxs-lookup"><span data-stu-id="f70e5-115">Modes</span></span> | <span data-ttu-id="f70e5-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-116">Return type</span></span> | <span data-ttu-id="f70e5-117">最低</span><span class="sxs-lookup"><span data-stu-id="f70e5-117">Minimum</span></span><br><span data-ttu-id="f70e5-118">要求集</span><span class="sxs-lookup"><span data-stu-id="f70e5-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f70e5-119">context</span><span class="sxs-lookup"><span data-stu-id="f70e5-119">context</span></span>](office.context.md) | <span data-ttu-id="f70e5-120">撰写</span><span class="sxs-lookup"><span data-stu-id="f70e5-120">Compose</span></span><br><span data-ttu-id="f70e5-121">读取</span><span class="sxs-lookup"><span data-stu-id="f70e5-121">Read</span></span> | [<span data-ttu-id="f70e5-122">Context</span><span class="sxs-lookup"><span data-stu-id="f70e5-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="f70e5-123">1.1</span><span class="sxs-lookup"><span data-stu-id="f70e5-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="f70e5-124">枚举</span><span class="sxs-lookup"><span data-stu-id="f70e5-124">Enumerations</span></span>

| <span data-ttu-id="f70e5-125">枚举</span><span class="sxs-lookup"><span data-stu-id="f70e5-125">Enumeration</span></span> | <span data-ttu-id="f70e5-126">型号</span><span class="sxs-lookup"><span data-stu-id="f70e5-126">Modes</span></span> | <span data-ttu-id="f70e5-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-127">Return type</span></span> | <span data-ttu-id="f70e5-128">最低</span><span class="sxs-lookup"><span data-stu-id="f70e5-128">Minimum</span></span><br><span data-ttu-id="f70e5-129">要求集</span><span class="sxs-lookup"><span data-stu-id="f70e5-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f70e5-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f70e5-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f70e5-131">撰写</span><span class="sxs-lookup"><span data-stu-id="f70e5-131">Compose</span></span><br><span data-ttu-id="f70e5-132">读取</span><span class="sxs-lookup"><span data-stu-id="f70e5-132">Read</span></span> | <span data-ttu-id="f70e5-133">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-133">String</span></span> | [<span data-ttu-id="f70e5-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f70e5-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f70e5-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f70e5-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f70e5-136">撰写</span><span class="sxs-lookup"><span data-stu-id="f70e5-136">Compose</span></span><br><span data-ttu-id="f70e5-137">读取</span><span class="sxs-lookup"><span data-stu-id="f70e5-137">Read</span></span> | <span data-ttu-id="f70e5-138">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-138">String</span></span> | [<span data-ttu-id="f70e5-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f70e5-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f70e5-140">EventType</span><span class="sxs-lookup"><span data-stu-id="f70e5-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f70e5-141">撰写</span><span class="sxs-lookup"><span data-stu-id="f70e5-141">Compose</span></span><br><span data-ttu-id="f70e5-142">读取</span><span class="sxs-lookup"><span data-stu-id="f70e5-142">Read</span></span> | <span data-ttu-id="f70e5-143">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-143">String</span></span> | [<span data-ttu-id="f70e5-144">1.5</span><span class="sxs-lookup"><span data-stu-id="f70e5-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f70e5-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f70e5-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f70e5-146">撰写</span><span class="sxs-lookup"><span data-stu-id="f70e5-146">Compose</span></span><br><span data-ttu-id="f70e5-147">读取</span><span class="sxs-lookup"><span data-stu-id="f70e5-147">Read</span></span> | <span data-ttu-id="f70e5-148">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-148">String</span></span> | [<span data-ttu-id="f70e5-149">1.1</span><span class="sxs-lookup"><span data-stu-id="f70e5-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="f70e5-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="f70e5-150">Namespaces</span></span>

<span data-ttu-id="f70e5-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="f70e5-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f70e5-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="f70e5-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f70e5-153">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="f70e5-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="f70e5-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="f70e5-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f70e5-155">类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-155">Type</span></span>

*   <span data-ttu-id="f70e5-156">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f70e5-157">属性：</span><span class="sxs-lookup"><span data-stu-id="f70e5-157">Properties:</span></span>

|<span data-ttu-id="f70e5-158">姓名</span><span class="sxs-lookup"><span data-stu-id="f70e5-158">Name</span></span>| <span data-ttu-id="f70e5-159">类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-159">Type</span></span>| <span data-ttu-id="f70e5-160">说明</span><span class="sxs-lookup"><span data-stu-id="f70e5-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f70e5-161">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-161">String</span></span>|<span data-ttu-id="f70e5-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="f70e5-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f70e5-163">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-163">String</span></span>|<span data-ttu-id="f70e5-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="f70e5-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70e5-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="f70e5-165">Requirements</span></span>

|<span data-ttu-id="f70e5-166">要求</span><span class="sxs-lookup"><span data-stu-id="f70e5-166">Requirement</span></span>| <span data-ttu-id="f70e5-167">值</span><span class="sxs-lookup"><span data-stu-id="f70e5-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70e5-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f70e5-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f70e5-169">1.1</span><span class="sxs-lookup"><span data-stu-id="f70e5-169">1.1</span></span>|
|[<span data-ttu-id="f70e5-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f70e5-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f70e5-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f70e5-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f70e5-172">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="f70e5-172">CoercionType: String</span></span>

<span data-ttu-id="f70e5-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="f70e5-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f70e5-174">类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-174">Type</span></span>

*   <span data-ttu-id="f70e5-175">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f70e5-176">属性：</span><span class="sxs-lookup"><span data-stu-id="f70e5-176">Properties:</span></span>

|<span data-ttu-id="f70e5-177">姓名</span><span class="sxs-lookup"><span data-stu-id="f70e5-177">Name</span></span>| <span data-ttu-id="f70e5-178">类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-178">Type</span></span>| <span data-ttu-id="f70e5-179">说明</span><span class="sxs-lookup"><span data-stu-id="f70e5-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f70e5-180">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-180">String</span></span>|<span data-ttu-id="f70e5-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="f70e5-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f70e5-182">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-182">String</span></span>|<span data-ttu-id="f70e5-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="f70e5-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70e5-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="f70e5-184">Requirements</span></span>

|<span data-ttu-id="f70e5-185">要求</span><span class="sxs-lookup"><span data-stu-id="f70e5-185">Requirement</span></span>| <span data-ttu-id="f70e5-186">值</span><span class="sxs-lookup"><span data-stu-id="f70e5-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70e5-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f70e5-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f70e5-188">1.1</span><span class="sxs-lookup"><span data-stu-id="f70e5-188">1.1</span></span>|
|[<span data-ttu-id="f70e5-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f70e5-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f70e5-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f70e5-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f70e5-191">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="f70e5-191">EventType: String</span></span>

<span data-ttu-id="f70e5-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="f70e5-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f70e5-193">类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-193">Type</span></span>

*   <span data-ttu-id="f70e5-194">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f70e5-195">属性：</span><span class="sxs-lookup"><span data-stu-id="f70e5-195">Properties:</span></span>

| <span data-ttu-id="f70e5-196">姓名</span><span class="sxs-lookup"><span data-stu-id="f70e5-196">Name</span></span> | <span data-ttu-id="f70e5-197">类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-197">Type</span></span> | <span data-ttu-id="f70e5-198">说明</span><span class="sxs-lookup"><span data-stu-id="f70e5-198">Description</span></span> | <span data-ttu-id="f70e5-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="f70e5-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="f70e5-200">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-200">String</span></span> | <span data-ttu-id="f70e5-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="f70e5-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f70e5-202">1.7</span><span class="sxs-lookup"><span data-stu-id="f70e5-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="f70e5-203">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-203">String</span></span> | <span data-ttu-id="f70e5-204">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="f70e5-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="f70e5-205">1.8</span><span class="sxs-lookup"><span data-stu-id="f70e5-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="f70e5-206">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-206">String</span></span> | <span data-ttu-id="f70e5-207">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="f70e5-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="f70e5-208">1.8</span><span class="sxs-lookup"><span data-stu-id="f70e5-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="f70e5-209">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-209">String</span></span> | <span data-ttu-id="f70e5-210">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="f70e5-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f70e5-211">1.5</span><span class="sxs-lookup"><span data-stu-id="f70e5-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="f70e5-212">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-212">String</span></span> | <span data-ttu-id="f70e5-213">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="f70e5-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="f70e5-214">预览</span><span class="sxs-lookup"><span data-stu-id="f70e5-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f70e5-215">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-215">String</span></span> | <span data-ttu-id="f70e5-216">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="f70e5-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f70e5-217">1.7</span><span class="sxs-lookup"><span data-stu-id="f70e5-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f70e5-218">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-218">String</span></span> | <span data-ttu-id="f70e5-219">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="f70e5-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f70e5-220">1.7</span><span class="sxs-lookup"><span data-stu-id="f70e5-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f70e5-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="f70e5-221">Requirements</span></span>

|<span data-ttu-id="f70e5-222">要求</span><span class="sxs-lookup"><span data-stu-id="f70e5-222">Requirement</span></span>| <span data-ttu-id="f70e5-223">值</span><span class="sxs-lookup"><span data-stu-id="f70e5-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70e5-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f70e5-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f70e5-225">1.5</span><span class="sxs-lookup"><span data-stu-id="f70e5-225">1.5</span></span> |
|[<span data-ttu-id="f70e5-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f70e5-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f70e5-227">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f70e5-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f70e5-228">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="f70e5-228">SourceProperty: String</span></span>

<span data-ttu-id="f70e5-229">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="f70e5-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f70e5-230">类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-230">Type</span></span>

*   <span data-ttu-id="f70e5-231">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f70e5-232">属性：</span><span class="sxs-lookup"><span data-stu-id="f70e5-232">Properties:</span></span>

|<span data-ttu-id="f70e5-233">姓名</span><span class="sxs-lookup"><span data-stu-id="f70e5-233">Name</span></span>| <span data-ttu-id="f70e5-234">类型</span><span class="sxs-lookup"><span data-stu-id="f70e5-234">Type</span></span>| <span data-ttu-id="f70e5-235">说明</span><span class="sxs-lookup"><span data-stu-id="f70e5-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f70e5-236">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-236">String</span></span>|<span data-ttu-id="f70e5-237">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="f70e5-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f70e5-238">String</span><span class="sxs-lookup"><span data-stu-id="f70e5-238">String</span></span>|<span data-ttu-id="f70e5-239">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="f70e5-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f70e5-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="f70e5-240">Requirements</span></span>

|<span data-ttu-id="f70e5-241">要求</span><span class="sxs-lookup"><span data-stu-id="f70e5-241">Requirement</span></span>| <span data-ttu-id="f70e5-242">值</span><span class="sxs-lookup"><span data-stu-id="f70e5-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="f70e5-243">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f70e5-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f70e5-244">1.1</span><span class="sxs-lookup"><span data-stu-id="f70e5-244">1.1</span></span>|
|[<span data-ttu-id="f70e5-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f70e5-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f70e5-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f70e5-246">Compose or Read</span></span>|
