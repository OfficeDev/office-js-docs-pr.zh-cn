---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ef9634058fcdc633e9ad3a0adb74c4abebf8038b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815058"
---
# <a name="office"></a><span data-ttu-id="5f241-102">Office</span><span class="sxs-lookup"><span data-stu-id="5f241-102">Office</span></span>

<span data-ttu-id="5f241-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="5f241-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f241-105">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-105">Requirements</span></span>

|<span data-ttu-id="5f241-106">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-106">Requirement</span></span>| <span data-ttu-id="5f241-107">值</span><span class="sxs-lookup"><span data-stu-id="5f241-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f241-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5f241-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5f241-109">1.1</span><span class="sxs-lookup"><span data-stu-id="5f241-109">1.1</span></span>|
|[<span data-ttu-id="5f241-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5f241-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5f241-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5f241-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="5f241-112">属性</span><span class="sxs-lookup"><span data-stu-id="5f241-112">Properties</span></span>

| <span data-ttu-id="5f241-113">属性</span><span class="sxs-lookup"><span data-stu-id="5f241-113">Property</span></span> | <span data-ttu-id="5f241-114">型号</span><span class="sxs-lookup"><span data-stu-id="5f241-114">Modes</span></span> | <span data-ttu-id="5f241-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="5f241-115">Return type</span></span> | <span data-ttu-id="5f241-116">最低</span><span class="sxs-lookup"><span data-stu-id="5f241-116">Minimum</span></span><br><span data-ttu-id="5f241-117">要求集</span><span class="sxs-lookup"><span data-stu-id="5f241-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5f241-118">context</span><span class="sxs-lookup"><span data-stu-id="5f241-118">context</span></span>](office.context.md) | <span data-ttu-id="5f241-119">撰写</span><span class="sxs-lookup"><span data-stu-id="5f241-119">Compose</span></span><br><span data-ttu-id="5f241-120">读取</span><span class="sxs-lookup"><span data-stu-id="5f241-120">Read</span></span> | [<span data-ttu-id="5f241-121">Context</span><span class="sxs-lookup"><span data-stu-id="5f241-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="5f241-122">1.1</span><span class="sxs-lookup"><span data-stu-id="5f241-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="5f241-123">枚举</span><span class="sxs-lookup"><span data-stu-id="5f241-123">Enumerations</span></span>

| <span data-ttu-id="5f241-124">枚举</span><span class="sxs-lookup"><span data-stu-id="5f241-124">Enumeration</span></span> | <span data-ttu-id="5f241-125">型号</span><span class="sxs-lookup"><span data-stu-id="5f241-125">Modes</span></span> | <span data-ttu-id="5f241-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="5f241-126">Return type</span></span> | <span data-ttu-id="5f241-127">最低</span><span class="sxs-lookup"><span data-stu-id="5f241-127">Minimum</span></span><br><span data-ttu-id="5f241-128">要求集</span><span class="sxs-lookup"><span data-stu-id="5f241-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5f241-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5f241-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5f241-130">撰写</span><span class="sxs-lookup"><span data-stu-id="5f241-130">Compose</span></span><br><span data-ttu-id="5f241-131">读取</span><span class="sxs-lookup"><span data-stu-id="5f241-131">Read</span></span> | <span data-ttu-id="5f241-132">String</span><span class="sxs-lookup"><span data-stu-id="5f241-132">String</span></span> | [<span data-ttu-id="5f241-133">1.1</span><span class="sxs-lookup"><span data-stu-id="5f241-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5f241-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5f241-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5f241-135">撰写</span><span class="sxs-lookup"><span data-stu-id="5f241-135">Compose</span></span><br><span data-ttu-id="5f241-136">读取</span><span class="sxs-lookup"><span data-stu-id="5f241-136">Read</span></span> | <span data-ttu-id="5f241-137">String</span><span class="sxs-lookup"><span data-stu-id="5f241-137">String</span></span> | [<span data-ttu-id="5f241-138">1.1</span><span class="sxs-lookup"><span data-stu-id="5f241-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5f241-139">EventType</span><span class="sxs-lookup"><span data-stu-id="5f241-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5f241-140">撰写</span><span class="sxs-lookup"><span data-stu-id="5f241-140">Compose</span></span><br><span data-ttu-id="5f241-141">读取</span><span class="sxs-lookup"><span data-stu-id="5f241-141">Read</span></span> | <span data-ttu-id="5f241-142">String</span><span class="sxs-lookup"><span data-stu-id="5f241-142">String</span></span> | [<span data-ttu-id="5f241-143">1.5</span><span class="sxs-lookup"><span data-stu-id="5f241-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="5f241-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5f241-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5f241-145">撰写</span><span class="sxs-lookup"><span data-stu-id="5f241-145">Compose</span></span><br><span data-ttu-id="5f241-146">读取</span><span class="sxs-lookup"><span data-stu-id="5f241-146">Read</span></span> | <span data-ttu-id="5f241-147">String</span><span class="sxs-lookup"><span data-stu-id="5f241-147">String</span></span> | [<span data-ttu-id="5f241-148">1.1</span><span class="sxs-lookup"><span data-stu-id="5f241-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="5f241-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="5f241-149">Namespaces</span></span>

<span data-ttu-id="5f241-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="5f241-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="5f241-151">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="5f241-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="5f241-152">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="5f241-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="5f241-153">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="5f241-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5f241-154">类型</span><span class="sxs-lookup"><span data-stu-id="5f241-154">Type</span></span>

*   <span data-ttu-id="5f241-155">String</span><span class="sxs-lookup"><span data-stu-id="5f241-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5f241-156">属性：</span><span class="sxs-lookup"><span data-stu-id="5f241-156">Properties:</span></span>

|<span data-ttu-id="5f241-157">名称</span><span class="sxs-lookup"><span data-stu-id="5f241-157">Name</span></span>| <span data-ttu-id="5f241-158">类型</span><span class="sxs-lookup"><span data-stu-id="5f241-158">Type</span></span>| <span data-ttu-id="5f241-159">说明</span><span class="sxs-lookup"><span data-stu-id="5f241-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5f241-160">String</span><span class="sxs-lookup"><span data-stu-id="5f241-160">String</span></span>|<span data-ttu-id="5f241-161">调用成功。</span><span class="sxs-lookup"><span data-stu-id="5f241-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5f241-162">String</span><span class="sxs-lookup"><span data-stu-id="5f241-162">String</span></span>|<span data-ttu-id="5f241-163">调用失败。</span><span class="sxs-lookup"><span data-stu-id="5f241-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f241-164">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-164">Requirements</span></span>

|<span data-ttu-id="5f241-165">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-165">Requirement</span></span>| <span data-ttu-id="5f241-166">值</span><span class="sxs-lookup"><span data-stu-id="5f241-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f241-167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5f241-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5f241-168">1.1</span><span class="sxs-lookup"><span data-stu-id="5f241-168">1.1</span></span>|
|[<span data-ttu-id="5f241-169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5f241-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5f241-170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5f241-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="5f241-171">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="5f241-171">CoercionType: String</span></span>

<span data-ttu-id="5f241-172">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="5f241-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5f241-173">类型</span><span class="sxs-lookup"><span data-stu-id="5f241-173">Type</span></span>

*   <span data-ttu-id="5f241-174">String</span><span class="sxs-lookup"><span data-stu-id="5f241-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5f241-175">属性：</span><span class="sxs-lookup"><span data-stu-id="5f241-175">Properties:</span></span>

|<span data-ttu-id="5f241-176">名称</span><span class="sxs-lookup"><span data-stu-id="5f241-176">Name</span></span>| <span data-ttu-id="5f241-177">类型</span><span class="sxs-lookup"><span data-stu-id="5f241-177">Type</span></span>| <span data-ttu-id="5f241-178">说明</span><span class="sxs-lookup"><span data-stu-id="5f241-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5f241-179">String</span><span class="sxs-lookup"><span data-stu-id="5f241-179">String</span></span>|<span data-ttu-id="5f241-180">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="5f241-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5f241-181">String</span><span class="sxs-lookup"><span data-stu-id="5f241-181">String</span></span>|<span data-ttu-id="5f241-182">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="5f241-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f241-183">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-183">Requirements</span></span>

|<span data-ttu-id="5f241-184">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-184">Requirement</span></span>| <span data-ttu-id="5f241-185">值</span><span class="sxs-lookup"><span data-stu-id="5f241-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f241-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5f241-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5f241-187">1.1</span><span class="sxs-lookup"><span data-stu-id="5f241-187">1.1</span></span>|
|[<span data-ttu-id="5f241-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5f241-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5f241-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5f241-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="5f241-190">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="5f241-190">EventType: String</span></span>

<span data-ttu-id="5f241-191">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="5f241-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5f241-192">类型</span><span class="sxs-lookup"><span data-stu-id="5f241-192">Type</span></span>

*   <span data-ttu-id="5f241-193">String</span><span class="sxs-lookup"><span data-stu-id="5f241-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5f241-194">属性：</span><span class="sxs-lookup"><span data-stu-id="5f241-194">Properties:</span></span>

| <span data-ttu-id="5f241-195">名称</span><span class="sxs-lookup"><span data-stu-id="5f241-195">Name</span></span> | <span data-ttu-id="5f241-196">类型</span><span class="sxs-lookup"><span data-stu-id="5f241-196">Type</span></span> | <span data-ttu-id="5f241-197">说明</span><span class="sxs-lookup"><span data-stu-id="5f241-197">Description</span></span> | <span data-ttu-id="5f241-198">最低要求集</span><span class="sxs-lookup"><span data-stu-id="5f241-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="5f241-199">String</span><span class="sxs-lookup"><span data-stu-id="5f241-199">String</span></span> | <span data-ttu-id="5f241-200">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="5f241-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5f241-201">1.7</span><span class="sxs-lookup"><span data-stu-id="5f241-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="5f241-202">String</span><span class="sxs-lookup"><span data-stu-id="5f241-202">String</span></span> | <span data-ttu-id="5f241-203">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="5f241-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="5f241-204">1.8</span><span class="sxs-lookup"><span data-stu-id="5f241-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="5f241-205">String</span><span class="sxs-lookup"><span data-stu-id="5f241-205">String</span></span> | <span data-ttu-id="5f241-206">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="5f241-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="5f241-207">1.8</span><span class="sxs-lookup"><span data-stu-id="5f241-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="5f241-208">String</span><span class="sxs-lookup"><span data-stu-id="5f241-208">String</span></span> | <span data-ttu-id="5f241-209">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="5f241-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5f241-210">1.5</span><span class="sxs-lookup"><span data-stu-id="5f241-210">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="5f241-211">String</span><span class="sxs-lookup"><span data-stu-id="5f241-211">String</span></span> | <span data-ttu-id="5f241-212">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="5f241-212">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="5f241-213">预览</span><span class="sxs-lookup"><span data-stu-id="5f241-213">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5f241-214">String</span><span class="sxs-lookup"><span data-stu-id="5f241-214">String</span></span> | <span data-ttu-id="5f241-215">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="5f241-215">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5f241-216">1.7</span><span class="sxs-lookup"><span data-stu-id="5f241-216">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5f241-217">String</span><span class="sxs-lookup"><span data-stu-id="5f241-217">String</span></span> | <span data-ttu-id="5f241-218">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="5f241-218">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5f241-219">1.7</span><span class="sxs-lookup"><span data-stu-id="5f241-219">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5f241-220">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-220">Requirements</span></span>

|<span data-ttu-id="5f241-221">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-221">Requirement</span></span>| <span data-ttu-id="5f241-222">值</span><span class="sxs-lookup"><span data-stu-id="5f241-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f241-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5f241-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5f241-224">1.5</span><span class="sxs-lookup"><span data-stu-id="5f241-224">1.5</span></span> |
|[<span data-ttu-id="5f241-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5f241-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5f241-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5f241-226">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="5f241-227">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="5f241-227">SourceProperty: String</span></span>

<span data-ttu-id="5f241-228">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="5f241-228">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5f241-229">类型</span><span class="sxs-lookup"><span data-stu-id="5f241-229">Type</span></span>

*   <span data-ttu-id="5f241-230">String</span><span class="sxs-lookup"><span data-stu-id="5f241-230">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5f241-231">属性：</span><span class="sxs-lookup"><span data-stu-id="5f241-231">Properties:</span></span>

|<span data-ttu-id="5f241-232">名称</span><span class="sxs-lookup"><span data-stu-id="5f241-232">Name</span></span>| <span data-ttu-id="5f241-233">类型</span><span class="sxs-lookup"><span data-stu-id="5f241-233">Type</span></span>| <span data-ttu-id="5f241-234">说明</span><span class="sxs-lookup"><span data-stu-id="5f241-234">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5f241-235">String</span><span class="sxs-lookup"><span data-stu-id="5f241-235">String</span></span>|<span data-ttu-id="5f241-236">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="5f241-236">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5f241-237">String</span><span class="sxs-lookup"><span data-stu-id="5f241-237">String</span></span>|<span data-ttu-id="5f241-238">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="5f241-238">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f241-239">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-239">Requirements</span></span>

|<span data-ttu-id="5f241-240">要求</span><span class="sxs-lookup"><span data-stu-id="5f241-240">Requirement</span></span>| <span data-ttu-id="5f241-241">值</span><span class="sxs-lookup"><span data-stu-id="5f241-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f241-242">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5f241-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5f241-243">1.1</span><span class="sxs-lookup"><span data-stu-id="5f241-243">1.1</span></span>|
|[<span data-ttu-id="5f241-244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5f241-244">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5f241-245">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5f241-245">Compose or Read</span></span>|
