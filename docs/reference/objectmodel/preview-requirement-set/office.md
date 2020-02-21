---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 2cd04cc6d333439a679803e39357e4d19c550f95
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165508"
---
# <a name="office"></a><span data-ttu-id="8ee8e-102">Office</span><span class="sxs-lookup"><span data-stu-id="8ee8e-102">Office</span></span>

<span data-ttu-id="8ee8e-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ee8e-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ee8e-105">Requirements</span></span>

|<span data-ttu-id="8ee8e-106">要求</span><span class="sxs-lookup"><span data-stu-id="8ee8e-106">Requirement</span></span>| <span data-ttu-id="8ee8e-107">值</span><span class="sxs-lookup"><span data-stu-id="8ee8e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ee8e-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8ee8e-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8ee8e-109">1.1</span><span class="sxs-lookup"><span data-stu-id="8ee8e-109">1.1</span></span>|
|[<span data-ttu-id="8ee8e-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8ee8e-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8ee8e-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8ee8e-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8ee8e-112">属性</span><span class="sxs-lookup"><span data-stu-id="8ee8e-112">Properties</span></span>

| <span data-ttu-id="8ee8e-113">属性</span><span class="sxs-lookup"><span data-stu-id="8ee8e-113">Property</span></span> | <span data-ttu-id="8ee8e-114">型号</span><span class="sxs-lookup"><span data-stu-id="8ee8e-114">Modes</span></span> | <span data-ttu-id="8ee8e-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-115">Return type</span></span> | <span data-ttu-id="8ee8e-116">最低</span><span class="sxs-lookup"><span data-stu-id="8ee8e-116">Minimum</span></span><br><span data-ttu-id="8ee8e-117">要求集</span><span class="sxs-lookup"><span data-stu-id="8ee8e-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8ee8e-118">context</span><span class="sxs-lookup"><span data-stu-id="8ee8e-118">context</span></span>](office.context.md) | <span data-ttu-id="8ee8e-119">撰写</span><span class="sxs-lookup"><span data-stu-id="8ee8e-119">Compose</span></span><br><span data-ttu-id="8ee8e-120">读取</span><span class="sxs-lookup"><span data-stu-id="8ee8e-120">Read</span></span> | [<span data-ttu-id="8ee8e-121">Context</span><span class="sxs-lookup"><span data-stu-id="8ee8e-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="8ee8e-122">1.1</span><span class="sxs-lookup"><span data-stu-id="8ee8e-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="8ee8e-123">枚举</span><span class="sxs-lookup"><span data-stu-id="8ee8e-123">Enumerations</span></span>

| <span data-ttu-id="8ee8e-124">枚举</span><span class="sxs-lookup"><span data-stu-id="8ee8e-124">Enumeration</span></span> | <span data-ttu-id="8ee8e-125">型号</span><span class="sxs-lookup"><span data-stu-id="8ee8e-125">Modes</span></span> | <span data-ttu-id="8ee8e-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-126">Return type</span></span> | <span data-ttu-id="8ee8e-127">最低</span><span class="sxs-lookup"><span data-stu-id="8ee8e-127">Minimum</span></span><br><span data-ttu-id="8ee8e-128">要求集</span><span class="sxs-lookup"><span data-stu-id="8ee8e-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8ee8e-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8ee8e-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8ee8e-130">撰写</span><span class="sxs-lookup"><span data-stu-id="8ee8e-130">Compose</span></span><br><span data-ttu-id="8ee8e-131">读取</span><span class="sxs-lookup"><span data-stu-id="8ee8e-131">Read</span></span> | <span data-ttu-id="8ee8e-132">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-132">String</span></span> | [<span data-ttu-id="8ee8e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="8ee8e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8ee8e-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8ee8e-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8ee8e-135">撰写</span><span class="sxs-lookup"><span data-stu-id="8ee8e-135">Compose</span></span><br><span data-ttu-id="8ee8e-136">读取</span><span class="sxs-lookup"><span data-stu-id="8ee8e-136">Read</span></span> | <span data-ttu-id="8ee8e-137">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-137">String</span></span> | [<span data-ttu-id="8ee8e-138">1.1</span><span class="sxs-lookup"><span data-stu-id="8ee8e-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8ee8e-139">EventType</span><span class="sxs-lookup"><span data-stu-id="8ee8e-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8ee8e-140">撰写</span><span class="sxs-lookup"><span data-stu-id="8ee8e-140">Compose</span></span><br><span data-ttu-id="8ee8e-141">读取</span><span class="sxs-lookup"><span data-stu-id="8ee8e-141">Read</span></span> | <span data-ttu-id="8ee8e-142">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-142">String</span></span> | [<span data-ttu-id="8ee8e-143">1.5</span><span class="sxs-lookup"><span data-stu-id="8ee8e-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="8ee8e-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8ee8e-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8ee8e-145">撰写</span><span class="sxs-lookup"><span data-stu-id="8ee8e-145">Compose</span></span><br><span data-ttu-id="8ee8e-146">读取</span><span class="sxs-lookup"><span data-stu-id="8ee8e-146">Read</span></span> | <span data-ttu-id="8ee8e-147">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-147">String</span></span> | [<span data-ttu-id="8ee8e-148">1.1</span><span class="sxs-lookup"><span data-stu-id="8ee8e-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="8ee8e-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="8ee8e-149">Namespaces</span></span>

<span data-ttu-id="8ee8e-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="8ee8e-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="8ee8e-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="8ee8e-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="8ee8e-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8ee8e-154">类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-154">Type</span></span>

*   <span data-ttu-id="8ee8e-155">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8ee8e-156">属性：</span><span class="sxs-lookup"><span data-stu-id="8ee8e-156">Properties:</span></span>

|<span data-ttu-id="8ee8e-157">名称</span><span class="sxs-lookup"><span data-stu-id="8ee8e-157">Name</span></span>| <span data-ttu-id="8ee8e-158">类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-158">Type</span></span>| <span data-ttu-id="8ee8e-159">说明</span><span class="sxs-lookup"><span data-stu-id="8ee8e-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8ee8e-160">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-160">String</span></span>|<span data-ttu-id="8ee8e-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8ee8e-162">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-162">String</span></span>|<span data-ttu-id="8ee8e-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ee8e-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ee8e-164">Requirements</span></span>

|<span data-ttu-id="8ee8e-165">要求</span><span class="sxs-lookup"><span data-stu-id="8ee8e-165">Requirement</span></span>| <span data-ttu-id="8ee8e-166">值</span><span class="sxs-lookup"><span data-stu-id="8ee8e-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ee8e-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8ee8e-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8ee8e-168">1.1</span><span class="sxs-lookup"><span data-stu-id="8ee8e-168">1.1</span></span>|
|[<span data-ttu-id="8ee8e-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8ee8e-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8ee8e-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8ee8e-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="8ee8e-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-171">CoercionType: String</span></span>

<span data-ttu-id="8ee8e-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8ee8e-173">类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-173">Type</span></span>

*   <span data-ttu-id="8ee8e-174">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8ee8e-175">属性：</span><span class="sxs-lookup"><span data-stu-id="8ee8e-175">Properties:</span></span>

|<span data-ttu-id="8ee8e-176">名称</span><span class="sxs-lookup"><span data-stu-id="8ee8e-176">Name</span></span>| <span data-ttu-id="8ee8e-177">类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-177">Type</span></span>| <span data-ttu-id="8ee8e-178">说明</span><span class="sxs-lookup"><span data-stu-id="8ee8e-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8ee8e-179">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-179">String</span></span>|<span data-ttu-id="8ee8e-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8ee8e-181">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-181">String</span></span>|<span data-ttu-id="8ee8e-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ee8e-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ee8e-183">Requirements</span></span>

|<span data-ttu-id="8ee8e-184">要求</span><span class="sxs-lookup"><span data-stu-id="8ee8e-184">Requirement</span></span>| <span data-ttu-id="8ee8e-185">值</span><span class="sxs-lookup"><span data-stu-id="8ee8e-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ee8e-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8ee8e-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8ee8e-187">1.1</span><span class="sxs-lookup"><span data-stu-id="8ee8e-187">1.1</span></span>|
|[<span data-ttu-id="8ee8e-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8ee8e-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8ee8e-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8ee8e-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="8ee8e-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-190">EventType: String</span></span>

<span data-ttu-id="8ee8e-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8ee8e-192">类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-192">Type</span></span>

*   <span data-ttu-id="8ee8e-193">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8ee8e-194">属性：</span><span class="sxs-lookup"><span data-stu-id="8ee8e-194">Properties:</span></span>

| <span data-ttu-id="8ee8e-195">名称</span><span class="sxs-lookup"><span data-stu-id="8ee8e-195">Name</span></span> | <span data-ttu-id="8ee8e-196">类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-196">Type</span></span> | <span data-ttu-id="8ee8e-197">说明</span><span class="sxs-lookup"><span data-stu-id="8ee8e-197">Description</span></span> | <span data-ttu-id="8ee8e-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="8ee8e-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="8ee8e-199">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-199">String</span></span> | <span data-ttu-id="8ee8e-200">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="8ee8e-201">1.7</span><span class="sxs-lookup"><span data-stu-id="8ee8e-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="8ee8e-202">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-202">String</span></span> | <span data-ttu-id="8ee8e-203">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="8ee8e-204">1.8</span><span class="sxs-lookup"><span data-stu-id="8ee8e-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="8ee8e-205">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-205">String</span></span> | <span data-ttu-id="8ee8e-206">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="8ee8e-207">1.8</span><span class="sxs-lookup"><span data-stu-id="8ee8e-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="8ee8e-208">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-208">String</span></span> | <span data-ttu-id="8ee8e-209">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="8ee8e-210">1.5</span><span class="sxs-lookup"><span data-stu-id="8ee8e-210">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="8ee8e-211">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-211">String</span></span> | <span data-ttu-id="8ee8e-212">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-212">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="8ee8e-213">预览</span><span class="sxs-lookup"><span data-stu-id="8ee8e-213">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="8ee8e-214">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-214">String</span></span> | <span data-ttu-id="8ee8e-215">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-215">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="8ee8e-216">1.7</span><span class="sxs-lookup"><span data-stu-id="8ee8e-216">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="8ee8e-217">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-217">String</span></span> | <span data-ttu-id="8ee8e-218">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-218">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="8ee8e-219">1.7</span><span class="sxs-lookup"><span data-stu-id="8ee8e-219">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ee8e-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ee8e-220">Requirements</span></span>

|<span data-ttu-id="8ee8e-221">要求</span><span class="sxs-lookup"><span data-stu-id="8ee8e-221">Requirement</span></span>| <span data-ttu-id="8ee8e-222">值</span><span class="sxs-lookup"><span data-stu-id="8ee8e-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ee8e-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8ee8e-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8ee8e-224">1.5</span><span class="sxs-lookup"><span data-stu-id="8ee8e-224">1.5</span></span> |
|[<span data-ttu-id="8ee8e-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8ee8e-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8ee8e-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8ee8e-226">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="8ee8e-227">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-227">SourceProperty: String</span></span>

<span data-ttu-id="8ee8e-228">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-228">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8ee8e-229">类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-229">Type</span></span>

*   <span data-ttu-id="8ee8e-230">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-230">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8ee8e-231">属性：</span><span class="sxs-lookup"><span data-stu-id="8ee8e-231">Properties:</span></span>

|<span data-ttu-id="8ee8e-232">名称</span><span class="sxs-lookup"><span data-stu-id="8ee8e-232">Name</span></span>| <span data-ttu-id="8ee8e-233">类型</span><span class="sxs-lookup"><span data-stu-id="8ee8e-233">Type</span></span>| <span data-ttu-id="8ee8e-234">说明</span><span class="sxs-lookup"><span data-stu-id="8ee8e-234">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8ee8e-235">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-235">String</span></span>|<span data-ttu-id="8ee8e-236">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-236">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8ee8e-237">String</span><span class="sxs-lookup"><span data-stu-id="8ee8e-237">String</span></span>|<span data-ttu-id="8ee8e-238">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="8ee8e-238">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ee8e-239">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ee8e-239">Requirements</span></span>

|<span data-ttu-id="8ee8e-240">要求</span><span class="sxs-lookup"><span data-stu-id="8ee8e-240">Requirement</span></span>| <span data-ttu-id="8ee8e-241">值</span><span class="sxs-lookup"><span data-stu-id="8ee8e-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ee8e-242">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8ee8e-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8ee8e-243">1.1</span><span class="sxs-lookup"><span data-stu-id="8ee8e-243">1.1</span></span>|
|[<span data-ttu-id="8ee8e-244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8ee8e-244">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8ee8e-245">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8ee8e-245">Compose or Read</span></span>|
