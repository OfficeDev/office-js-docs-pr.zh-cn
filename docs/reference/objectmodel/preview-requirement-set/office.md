---
title: Office 命名空间 - 预览要求集
description: 使用邮箱 API preview 要求集的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 634b8593e1d1a58b61c4a330ed96611903e4a27e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611608"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="4897d-103">Office （邮箱预览要求集）</span><span class="sxs-lookup"><span data-stu-id="4897d-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="4897d-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="4897d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4897d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4897d-106">Requirements</span></span>

|<span data-ttu-id="4897d-107">要求</span><span class="sxs-lookup"><span data-stu-id="4897d-107">Requirement</span></span>| <span data-ttu-id="4897d-108">值</span><span class="sxs-lookup"><span data-stu-id="4897d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4897d-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4897d-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4897d-110">1.1</span><span class="sxs-lookup"><span data-stu-id="4897d-110">1.1</span></span>|
|[<span data-ttu-id="4897d-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4897d-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4897d-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4897d-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="4897d-113">属性</span><span class="sxs-lookup"><span data-stu-id="4897d-113">Properties</span></span>

| <span data-ttu-id="4897d-114">属性</span><span class="sxs-lookup"><span data-stu-id="4897d-114">Property</span></span> | <span data-ttu-id="4897d-115">型号</span><span class="sxs-lookup"><span data-stu-id="4897d-115">Modes</span></span> | <span data-ttu-id="4897d-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="4897d-116">Return type</span></span> | <span data-ttu-id="4897d-117">最低</span><span class="sxs-lookup"><span data-stu-id="4897d-117">Minimum</span></span><br><span data-ttu-id="4897d-118">要求集</span><span class="sxs-lookup"><span data-stu-id="4897d-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4897d-119">context</span><span class="sxs-lookup"><span data-stu-id="4897d-119">context</span></span>](office.context.md) | <span data-ttu-id="4897d-120">撰写</span><span class="sxs-lookup"><span data-stu-id="4897d-120">Compose</span></span><br><span data-ttu-id="4897d-121">Read</span><span class="sxs-lookup"><span data-stu-id="4897d-121">Read</span></span> | [<span data-ttu-id="4897d-122">Context</span><span class="sxs-lookup"><span data-stu-id="4897d-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="4897d-123">1.1</span><span class="sxs-lookup"><span data-stu-id="4897d-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="4897d-124">枚举</span><span class="sxs-lookup"><span data-stu-id="4897d-124">Enumerations</span></span>

| <span data-ttu-id="4897d-125">枚举</span><span class="sxs-lookup"><span data-stu-id="4897d-125">Enumeration</span></span> | <span data-ttu-id="4897d-126">型号</span><span class="sxs-lookup"><span data-stu-id="4897d-126">Modes</span></span> | <span data-ttu-id="4897d-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="4897d-127">Return type</span></span> | <span data-ttu-id="4897d-128">最低</span><span class="sxs-lookup"><span data-stu-id="4897d-128">Minimum</span></span><br><span data-ttu-id="4897d-129">要求集</span><span class="sxs-lookup"><span data-stu-id="4897d-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4897d-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4897d-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4897d-131">撰写</span><span class="sxs-lookup"><span data-stu-id="4897d-131">Compose</span></span><br><span data-ttu-id="4897d-132">Read</span><span class="sxs-lookup"><span data-stu-id="4897d-132">Read</span></span> | <span data-ttu-id="4897d-133">String</span><span class="sxs-lookup"><span data-stu-id="4897d-133">String</span></span> | [<span data-ttu-id="4897d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="4897d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4897d-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4897d-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4897d-136">撰写</span><span class="sxs-lookup"><span data-stu-id="4897d-136">Compose</span></span><br><span data-ttu-id="4897d-137">Read</span><span class="sxs-lookup"><span data-stu-id="4897d-137">Read</span></span> | <span data-ttu-id="4897d-138">String</span><span class="sxs-lookup"><span data-stu-id="4897d-138">String</span></span> | [<span data-ttu-id="4897d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="4897d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4897d-140">EventType</span><span class="sxs-lookup"><span data-stu-id="4897d-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4897d-141">撰写</span><span class="sxs-lookup"><span data-stu-id="4897d-141">Compose</span></span><br><span data-ttu-id="4897d-142">Read</span><span class="sxs-lookup"><span data-stu-id="4897d-142">Read</span></span> | <span data-ttu-id="4897d-143">String</span><span class="sxs-lookup"><span data-stu-id="4897d-143">String</span></span> | [<span data-ttu-id="4897d-144">1.5</span><span class="sxs-lookup"><span data-stu-id="4897d-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="4897d-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4897d-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4897d-146">撰写</span><span class="sxs-lookup"><span data-stu-id="4897d-146">Compose</span></span><br><span data-ttu-id="4897d-147">Read</span><span class="sxs-lookup"><span data-stu-id="4897d-147">Read</span></span> | <span data-ttu-id="4897d-148">String</span><span class="sxs-lookup"><span data-stu-id="4897d-148">String</span></span> | [<span data-ttu-id="4897d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="4897d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="4897d-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="4897d-150">Namespaces</span></span>

<span data-ttu-id="4897d-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview)：包含许多特定于 Outlook 的枚举，例如、、、、、 `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` 和 `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="4897d-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="4897d-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="4897d-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4897d-153">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="4897d-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="4897d-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="4897d-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4897d-155">类型</span><span class="sxs-lookup"><span data-stu-id="4897d-155">Type</span></span>

*   <span data-ttu-id="4897d-156">String</span><span class="sxs-lookup"><span data-stu-id="4897d-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4897d-157">属性：</span><span class="sxs-lookup"><span data-stu-id="4897d-157">Properties:</span></span>

|<span data-ttu-id="4897d-158">名称</span><span class="sxs-lookup"><span data-stu-id="4897d-158">Name</span></span>| <span data-ttu-id="4897d-159">类型</span><span class="sxs-lookup"><span data-stu-id="4897d-159">Type</span></span>| <span data-ttu-id="4897d-160">说明</span><span class="sxs-lookup"><span data-stu-id="4897d-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4897d-161">String</span><span class="sxs-lookup"><span data-stu-id="4897d-161">String</span></span>|<span data-ttu-id="4897d-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="4897d-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4897d-163">String</span><span class="sxs-lookup"><span data-stu-id="4897d-163">String</span></span>|<span data-ttu-id="4897d-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="4897d-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4897d-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="4897d-165">Requirements</span></span>

|<span data-ttu-id="4897d-166">要求</span><span class="sxs-lookup"><span data-stu-id="4897d-166">Requirement</span></span>| <span data-ttu-id="4897d-167">值</span><span class="sxs-lookup"><span data-stu-id="4897d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4897d-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4897d-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4897d-169">1.1</span><span class="sxs-lookup"><span data-stu-id="4897d-169">1.1</span></span>|
|[<span data-ttu-id="4897d-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4897d-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4897d-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4897d-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="4897d-172">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="4897d-172">CoercionType: String</span></span>

<span data-ttu-id="4897d-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="4897d-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4897d-174">类型</span><span class="sxs-lookup"><span data-stu-id="4897d-174">Type</span></span>

*   <span data-ttu-id="4897d-175">String</span><span class="sxs-lookup"><span data-stu-id="4897d-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4897d-176">属性：</span><span class="sxs-lookup"><span data-stu-id="4897d-176">Properties:</span></span>

|<span data-ttu-id="4897d-177">名称</span><span class="sxs-lookup"><span data-stu-id="4897d-177">Name</span></span>| <span data-ttu-id="4897d-178">类型</span><span class="sxs-lookup"><span data-stu-id="4897d-178">Type</span></span>| <span data-ttu-id="4897d-179">说明</span><span class="sxs-lookup"><span data-stu-id="4897d-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4897d-180">String</span><span class="sxs-lookup"><span data-stu-id="4897d-180">String</span></span>|<span data-ttu-id="4897d-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="4897d-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4897d-182">String</span><span class="sxs-lookup"><span data-stu-id="4897d-182">String</span></span>|<span data-ttu-id="4897d-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="4897d-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4897d-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="4897d-184">Requirements</span></span>

|<span data-ttu-id="4897d-185">要求</span><span class="sxs-lookup"><span data-stu-id="4897d-185">Requirement</span></span>| <span data-ttu-id="4897d-186">值</span><span class="sxs-lookup"><span data-stu-id="4897d-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="4897d-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4897d-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4897d-188">1.1</span><span class="sxs-lookup"><span data-stu-id="4897d-188">1.1</span></span>|
|[<span data-ttu-id="4897d-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4897d-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4897d-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4897d-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="4897d-191">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="4897d-191">EventType: String</span></span>

<span data-ttu-id="4897d-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="4897d-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4897d-193">类型</span><span class="sxs-lookup"><span data-stu-id="4897d-193">Type</span></span>

*   <span data-ttu-id="4897d-194">String</span><span class="sxs-lookup"><span data-stu-id="4897d-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4897d-195">属性：</span><span class="sxs-lookup"><span data-stu-id="4897d-195">Properties:</span></span>

| <span data-ttu-id="4897d-196">名称</span><span class="sxs-lookup"><span data-stu-id="4897d-196">Name</span></span> | <span data-ttu-id="4897d-197">类型</span><span class="sxs-lookup"><span data-stu-id="4897d-197">Type</span></span> | <span data-ttu-id="4897d-198">Description</span><span class="sxs-lookup"><span data-stu-id="4897d-198">Description</span></span> | <span data-ttu-id="4897d-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="4897d-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="4897d-200">String</span><span class="sxs-lookup"><span data-stu-id="4897d-200">String</span></span> | <span data-ttu-id="4897d-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="4897d-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4897d-202">1.7</span><span class="sxs-lookup"><span data-stu-id="4897d-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="4897d-203">String</span><span class="sxs-lookup"><span data-stu-id="4897d-203">String</span></span> | <span data-ttu-id="4897d-204">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="4897d-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="4897d-205">1.8</span><span class="sxs-lookup"><span data-stu-id="4897d-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="4897d-206">String</span><span class="sxs-lookup"><span data-stu-id="4897d-206">String</span></span> | <span data-ttu-id="4897d-207">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="4897d-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="4897d-208">1.8</span><span class="sxs-lookup"><span data-stu-id="4897d-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="4897d-209">String</span><span class="sxs-lookup"><span data-stu-id="4897d-209">String</span></span> | <span data-ttu-id="4897d-210">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="4897d-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="4897d-211">1.5</span><span class="sxs-lookup"><span data-stu-id="4897d-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="4897d-212">String</span><span class="sxs-lookup"><span data-stu-id="4897d-212">String</span></span> | <span data-ttu-id="4897d-213">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="4897d-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="4897d-214">预览</span><span class="sxs-lookup"><span data-stu-id="4897d-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4897d-215">String</span><span class="sxs-lookup"><span data-stu-id="4897d-215">String</span></span> | <span data-ttu-id="4897d-216">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="4897d-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4897d-217">1.7</span><span class="sxs-lookup"><span data-stu-id="4897d-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4897d-218">String</span><span class="sxs-lookup"><span data-stu-id="4897d-218">String</span></span> | <span data-ttu-id="4897d-219">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="4897d-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4897d-220">1.7</span><span class="sxs-lookup"><span data-stu-id="4897d-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4897d-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="4897d-221">Requirements</span></span>

|<span data-ttu-id="4897d-222">要求</span><span class="sxs-lookup"><span data-stu-id="4897d-222">Requirement</span></span>| <span data-ttu-id="4897d-223">值</span><span class="sxs-lookup"><span data-stu-id="4897d-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="4897d-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4897d-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4897d-225">1.5</span><span class="sxs-lookup"><span data-stu-id="4897d-225">1.5</span></span> |
|[<span data-ttu-id="4897d-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4897d-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4897d-227">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4897d-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="4897d-228">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="4897d-228">SourceProperty: String</span></span>

<span data-ttu-id="4897d-229">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="4897d-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4897d-230">类型</span><span class="sxs-lookup"><span data-stu-id="4897d-230">Type</span></span>

*   <span data-ttu-id="4897d-231">String</span><span class="sxs-lookup"><span data-stu-id="4897d-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4897d-232">属性：</span><span class="sxs-lookup"><span data-stu-id="4897d-232">Properties:</span></span>

|<span data-ttu-id="4897d-233">名称</span><span class="sxs-lookup"><span data-stu-id="4897d-233">Name</span></span>| <span data-ttu-id="4897d-234">类型</span><span class="sxs-lookup"><span data-stu-id="4897d-234">Type</span></span>| <span data-ttu-id="4897d-235">说明</span><span class="sxs-lookup"><span data-stu-id="4897d-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4897d-236">String</span><span class="sxs-lookup"><span data-stu-id="4897d-236">String</span></span>|<span data-ttu-id="4897d-237">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="4897d-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4897d-238">String</span><span class="sxs-lookup"><span data-stu-id="4897d-238">String</span></span>|<span data-ttu-id="4897d-239">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="4897d-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4897d-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="4897d-240">Requirements</span></span>

|<span data-ttu-id="4897d-241">要求</span><span class="sxs-lookup"><span data-stu-id="4897d-241">Requirement</span></span>| <span data-ttu-id="4897d-242">值</span><span class="sxs-lookup"><span data-stu-id="4897d-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="4897d-243">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4897d-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4897d-244">1.1</span><span class="sxs-lookup"><span data-stu-id="4897d-244">1.1</span></span>|
|[<span data-ttu-id="4897d-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4897d-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4897d-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4897d-246">Compose or Read</span></span>|
