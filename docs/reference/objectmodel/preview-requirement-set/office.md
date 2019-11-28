---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: bd37b1be4d77d73cb56b0b2593ccc57dea6cab27
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629228"
---
# <a name="office"></a><span data-ttu-id="a2be0-102">Office</span><span class="sxs-lookup"><span data-stu-id="a2be0-102">Office</span></span>

<span data-ttu-id="a2be0-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="a2be0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a2be0-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2be0-105">Requirements</span></span>

|<span data-ttu-id="a2be0-106">要求</span><span class="sxs-lookup"><span data-stu-id="a2be0-106">Requirement</span></span>| <span data-ttu-id="a2be0-107">值</span><span class="sxs-lookup"><span data-stu-id="a2be0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2be0-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a2be0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2be0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a2be0-109">1.0</span></span>|
|[<span data-ttu-id="a2be0-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a2be0-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2be0-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a2be0-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a2be0-112">属性</span><span class="sxs-lookup"><span data-stu-id="a2be0-112">Properties</span></span>

| <span data-ttu-id="a2be0-113">属性</span><span class="sxs-lookup"><span data-stu-id="a2be0-113">Property</span></span> | <span data-ttu-id="a2be0-114">型号</span><span class="sxs-lookup"><span data-stu-id="a2be0-114">Modes</span></span> | <span data-ttu-id="a2be0-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-115">Return type</span></span> | <span data-ttu-id="a2be0-116">最低</span><span class="sxs-lookup"><span data-stu-id="a2be0-116">Minimum</span></span><br><span data-ttu-id="a2be0-117">要求集</span><span class="sxs-lookup"><span data-stu-id="a2be0-117">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="a2be0-118">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a2be0-118">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a2be0-119">撰写</span><span class="sxs-lookup"><span data-stu-id="a2be0-119">Compose</span></span><br><span data-ttu-id="a2be0-120">读取</span><span class="sxs-lookup"><span data-stu-id="a2be0-120">Read</span></span> | <span data-ttu-id="a2be0-121">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-121">String</span></span> | <span data-ttu-id="a2be0-122">1.0</span><span class="sxs-lookup"><span data-stu-id="a2be0-122">1.0</span></span> |
| [<span data-ttu-id="a2be0-123">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a2be0-123">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a2be0-124">撰写</span><span class="sxs-lookup"><span data-stu-id="a2be0-124">Compose</span></span><br><span data-ttu-id="a2be0-125">读取</span><span class="sxs-lookup"><span data-stu-id="a2be0-125">Read</span></span> | <span data-ttu-id="a2be0-126">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-126">String</span></span> | <span data-ttu-id="a2be0-127">1.0</span><span class="sxs-lookup"><span data-stu-id="a2be0-127">1.0</span></span> |
| [<span data-ttu-id="a2be0-128">EventType</span><span class="sxs-lookup"><span data-stu-id="a2be0-128">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a2be0-129">撰写</span><span class="sxs-lookup"><span data-stu-id="a2be0-129">Compose</span></span><br><span data-ttu-id="a2be0-130">读取</span><span class="sxs-lookup"><span data-stu-id="a2be0-130">Read</span></span> | <span data-ttu-id="a2be0-131">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-131">String</span></span> | <span data-ttu-id="a2be0-132">1.5</span><span class="sxs-lookup"><span data-stu-id="a2be0-132">1.5</span></span> |
| [<span data-ttu-id="a2be0-133">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a2be0-133">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a2be0-134">撰写</span><span class="sxs-lookup"><span data-stu-id="a2be0-134">Compose</span></span><br><span data-ttu-id="a2be0-135">读取</span><span class="sxs-lookup"><span data-stu-id="a2be0-135">Read</span></span> | <span data-ttu-id="a2be0-136">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-136">String</span></span> | <span data-ttu-id="a2be0-137">1.0</span><span class="sxs-lookup"><span data-stu-id="a2be0-137">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a2be0-138">命名空间</span><span class="sxs-lookup"><span data-stu-id="a2be0-138">Namespaces</span></span>

<span data-ttu-id="a2be0-139">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="a2be0-139">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a2be0-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)：包含多个`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="a2be0-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="property-details"></a><span data-ttu-id="a2be0-141">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="a2be0-141">Property details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a2be0-142">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="a2be0-142">AsyncResultStatus: String</span></span>

<span data-ttu-id="a2be0-143">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="a2be0-143">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a2be0-144">类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-144">Type</span></span>

*   <span data-ttu-id="a2be0-145">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-145">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2be0-146">属性：</span><span class="sxs-lookup"><span data-stu-id="a2be0-146">Properties:</span></span>

|<span data-ttu-id="a2be0-147">名称</span><span class="sxs-lookup"><span data-stu-id="a2be0-147">Name</span></span>| <span data-ttu-id="a2be0-148">类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-148">Type</span></span>| <span data-ttu-id="a2be0-149">说明</span><span class="sxs-lookup"><span data-stu-id="a2be0-149">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a2be0-150">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-150">String</span></span>|<span data-ttu-id="a2be0-151">调用成功。</span><span class="sxs-lookup"><span data-stu-id="a2be0-151">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a2be0-152">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-152">String</span></span>|<span data-ttu-id="a2be0-153">调用失败。</span><span class="sxs-lookup"><span data-stu-id="a2be0-153">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2be0-154">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2be0-154">Requirements</span></span>

|<span data-ttu-id="a2be0-155">要求</span><span class="sxs-lookup"><span data-stu-id="a2be0-155">Requirement</span></span>| <span data-ttu-id="a2be0-156">值</span><span class="sxs-lookup"><span data-stu-id="a2be0-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2be0-157">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a2be0-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2be0-158">1.0</span><span class="sxs-lookup"><span data-stu-id="a2be0-158">1.0</span></span>|
|[<span data-ttu-id="a2be0-159">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a2be0-159">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2be0-160">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a2be0-160">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="a2be0-161">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="a2be0-161">CoercionType: String</span></span>

<span data-ttu-id="a2be0-162">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="a2be0-162">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a2be0-163">类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-163">Type</span></span>

*   <span data-ttu-id="a2be0-164">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-164">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2be0-165">属性：</span><span class="sxs-lookup"><span data-stu-id="a2be0-165">Properties:</span></span>

|<span data-ttu-id="a2be0-166">名称</span><span class="sxs-lookup"><span data-stu-id="a2be0-166">Name</span></span>| <span data-ttu-id="a2be0-167">类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-167">Type</span></span>| <span data-ttu-id="a2be0-168">说明</span><span class="sxs-lookup"><span data-stu-id="a2be0-168">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a2be0-169">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-169">String</span></span>|<span data-ttu-id="a2be0-170">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="a2be0-170">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a2be0-171">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-171">String</span></span>|<span data-ttu-id="a2be0-172">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="a2be0-172">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2be0-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2be0-173">Requirements</span></span>

|<span data-ttu-id="a2be0-174">要求</span><span class="sxs-lookup"><span data-stu-id="a2be0-174">Requirement</span></span>| <span data-ttu-id="a2be0-175">值</span><span class="sxs-lookup"><span data-stu-id="a2be0-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2be0-176">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a2be0-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2be0-177">1.0</span><span class="sxs-lookup"><span data-stu-id="a2be0-177">1.0</span></span>|
|[<span data-ttu-id="a2be0-178">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a2be0-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2be0-179">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a2be0-179">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="a2be0-180">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="a2be0-180">EventType: String</span></span>

<span data-ttu-id="a2be0-181">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="a2be0-181">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a2be0-182">类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-182">Type</span></span>

*   <span data-ttu-id="a2be0-183">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-183">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2be0-184">属性：</span><span class="sxs-lookup"><span data-stu-id="a2be0-184">Properties:</span></span>

| <span data-ttu-id="a2be0-185">名称</span><span class="sxs-lookup"><span data-stu-id="a2be0-185">Name</span></span> | <span data-ttu-id="a2be0-186">类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-186">Type</span></span> | <span data-ttu-id="a2be0-187">说明</span><span class="sxs-lookup"><span data-stu-id="a2be0-187">Description</span></span> | <span data-ttu-id="a2be0-188">最低要求集</span><span class="sxs-lookup"><span data-stu-id="a2be0-188">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="a2be0-189">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-189">String</span></span> | <span data-ttu-id="a2be0-190">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="a2be0-190">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="a2be0-191">1.7</span><span class="sxs-lookup"><span data-stu-id="a2be0-191">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="a2be0-192">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-192">String</span></span> | <span data-ttu-id="a2be0-193">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="a2be0-193">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="a2be0-194">1.8</span><span class="sxs-lookup"><span data-stu-id="a2be0-194">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="a2be0-195">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-195">String</span></span> | <span data-ttu-id="a2be0-196">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="a2be0-196">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="a2be0-197">1.8</span><span class="sxs-lookup"><span data-stu-id="a2be0-197">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="a2be0-198">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-198">String</span></span> | <span data-ttu-id="a2be0-199">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="a2be0-199">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="a2be0-200">1.5</span><span class="sxs-lookup"><span data-stu-id="a2be0-200">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="a2be0-201">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-201">String</span></span> | <span data-ttu-id="a2be0-202">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="a2be0-202">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="a2be0-203">预览</span><span class="sxs-lookup"><span data-stu-id="a2be0-203">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="a2be0-204">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-204">String</span></span> | <span data-ttu-id="a2be0-205">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="a2be0-205">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="a2be0-206">1.7</span><span class="sxs-lookup"><span data-stu-id="a2be0-206">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="a2be0-207">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-207">String</span></span> | <span data-ttu-id="a2be0-208">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="a2be0-208">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="a2be0-209">1.7</span><span class="sxs-lookup"><span data-stu-id="a2be0-209">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a2be0-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2be0-210">Requirements</span></span>

|<span data-ttu-id="a2be0-211">要求</span><span class="sxs-lookup"><span data-stu-id="a2be0-211">Requirement</span></span>| <span data-ttu-id="a2be0-212">值</span><span class="sxs-lookup"><span data-stu-id="a2be0-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2be0-213">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a2be0-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2be0-214">1.5</span><span class="sxs-lookup"><span data-stu-id="a2be0-214">1.5</span></span> |
|[<span data-ttu-id="a2be0-215">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a2be0-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2be0-216">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a2be0-216">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a2be0-217">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="a2be0-217">SourceProperty: String</span></span>

<span data-ttu-id="a2be0-218">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="a2be0-218">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a2be0-219">类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-219">Type</span></span>

*   <span data-ttu-id="a2be0-220">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-220">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2be0-221">属性：</span><span class="sxs-lookup"><span data-stu-id="a2be0-221">Properties:</span></span>

|<span data-ttu-id="a2be0-222">名称</span><span class="sxs-lookup"><span data-stu-id="a2be0-222">Name</span></span>| <span data-ttu-id="a2be0-223">类型</span><span class="sxs-lookup"><span data-stu-id="a2be0-223">Type</span></span>| <span data-ttu-id="a2be0-224">说明</span><span class="sxs-lookup"><span data-stu-id="a2be0-224">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a2be0-225">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-225">String</span></span>|<span data-ttu-id="a2be0-226">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="a2be0-226">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a2be0-227">String</span><span class="sxs-lookup"><span data-stu-id="a2be0-227">String</span></span>|<span data-ttu-id="a2be0-228">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="a2be0-228">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2be0-229">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2be0-229">Requirements</span></span>

|<span data-ttu-id="a2be0-230">要求</span><span class="sxs-lookup"><span data-stu-id="a2be0-230">Requirement</span></span>| <span data-ttu-id="a2be0-231">值</span><span class="sxs-lookup"><span data-stu-id="a2be0-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2be0-232">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="a2be0-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2be0-233">1.0</span><span class="sxs-lookup"><span data-stu-id="a2be0-233">1.0</span></span>|
|[<span data-ttu-id="a2be0-234">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a2be0-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2be0-235">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a2be0-235">Compose or Read</span></span>|
