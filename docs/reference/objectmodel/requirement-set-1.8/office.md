---
title: Office 命名空间-要求集1。8
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 91a0bef2a8280a068763c98b17644bd9268e2fb4
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902146"
---
# <a name="office"></a><span data-ttu-id="c0c87-102">Office</span><span class="sxs-lookup"><span data-stu-id="c0c87-102">Office</span></span>

<span data-ttu-id="c0c87-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="c0c87-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0c87-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="c0c87-105">Requirements</span></span>

|<span data-ttu-id="c0c87-106">要求</span><span class="sxs-lookup"><span data-stu-id="c0c87-106">Requirement</span></span>| <span data-ttu-id="c0c87-107">值</span><span class="sxs-lookup"><span data-stu-id="c0c87-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c87-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c87-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c87-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c0c87-109">1.0</span></span>|
|[<span data-ttu-id="c0c87-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c87-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c87-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c87-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c0c87-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c0c87-112">Members and methods</span></span>

| <span data-ttu-id="c0c87-113">成员</span><span class="sxs-lookup"><span data-stu-id="c0c87-113">Member</span></span> | <span data-ttu-id="c0c87-114">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c0c87-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c0c87-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c0c87-116">Member</span><span class="sxs-lookup"><span data-stu-id="c0c87-116">Member</span></span> |
| [<span data-ttu-id="c0c87-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c0c87-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c0c87-118">Member</span><span class="sxs-lookup"><span data-stu-id="c0c87-118">Member</span></span> |
| [<span data-ttu-id="c0c87-119">EventType</span><span class="sxs-lookup"><span data-stu-id="c0c87-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c0c87-120">Member</span><span class="sxs-lookup"><span data-stu-id="c0c87-120">Member</span></span> |
| [<span data-ttu-id="c0c87-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c0c87-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c0c87-122">成员</span><span class="sxs-lookup"><span data-stu-id="c0c87-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c0c87-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="c0c87-123">Namespaces</span></span>

<span data-ttu-id="c0c87-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="c0c87-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c0c87-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8)：包含多个`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="c0c87-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="c0c87-126">Members</span><span class="sxs-lookup"><span data-stu-id="c0c87-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="c0c87-127">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="c0c87-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="c0c87-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="c0c87-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c0c87-129">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-129">Type</span></span>

*   <span data-ttu-id="c0c87-130">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c0c87-131">属性：</span><span class="sxs-lookup"><span data-stu-id="c0c87-131">Properties:</span></span>

|<span data-ttu-id="c0c87-132">名称</span><span class="sxs-lookup"><span data-stu-id="c0c87-132">Name</span></span>| <span data-ttu-id="c0c87-133">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-133">Type</span></span>| <span data-ttu-id="c0c87-134">说明</span><span class="sxs-lookup"><span data-stu-id="c0c87-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c0c87-135">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-135">String</span></span>|<span data-ttu-id="c0c87-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="c0c87-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c0c87-137">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-137">String</span></span>|<span data-ttu-id="c0c87-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="c0c87-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c0c87-139">Requirements</span><span class="sxs-lookup"><span data-stu-id="c0c87-139">Requirements</span></span>

|<span data-ttu-id="c0c87-140">要求</span><span class="sxs-lookup"><span data-stu-id="c0c87-140">Requirement</span></span>| <span data-ttu-id="c0c87-141">值</span><span class="sxs-lookup"><span data-stu-id="c0c87-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c87-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c87-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c87-143">1.0</span><span class="sxs-lookup"><span data-stu-id="c0c87-143">1.0</span></span>|
|[<span data-ttu-id="c0c87-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c87-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c87-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c87-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="c0c87-146">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="c0c87-146">CoercionType: String</span></span>

<span data-ttu-id="c0c87-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="c0c87-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c0c87-148">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-148">Type</span></span>

*   <span data-ttu-id="c0c87-149">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c0c87-150">属性：</span><span class="sxs-lookup"><span data-stu-id="c0c87-150">Properties:</span></span>

|<span data-ttu-id="c0c87-151">名称</span><span class="sxs-lookup"><span data-stu-id="c0c87-151">Name</span></span>| <span data-ttu-id="c0c87-152">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-152">Type</span></span>| <span data-ttu-id="c0c87-153">说明</span><span class="sxs-lookup"><span data-stu-id="c0c87-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c0c87-154">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-154">String</span></span>|<span data-ttu-id="c0c87-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="c0c87-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c0c87-156">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-156">String</span></span>|<span data-ttu-id="c0c87-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="c0c87-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c0c87-158">Requirements</span><span class="sxs-lookup"><span data-stu-id="c0c87-158">Requirements</span></span>

|<span data-ttu-id="c0c87-159">要求</span><span class="sxs-lookup"><span data-stu-id="c0c87-159">Requirement</span></span>| <span data-ttu-id="c0c87-160">值</span><span class="sxs-lookup"><span data-stu-id="c0c87-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c87-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c87-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c87-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c0c87-162">1.0</span></span>|
|[<span data-ttu-id="c0c87-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c87-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c87-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c87-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="c0c87-165">事件类型： String</span><span class="sxs-lookup"><span data-stu-id="c0c87-165">EventType: String</span></span>

<span data-ttu-id="c0c87-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="c0c87-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c0c87-167">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-167">Type</span></span>

*   <span data-ttu-id="c0c87-168">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c0c87-169">属性：</span><span class="sxs-lookup"><span data-stu-id="c0c87-169">Properties:</span></span>

| <span data-ttu-id="c0c87-170">名称</span><span class="sxs-lookup"><span data-stu-id="c0c87-170">Name</span></span> | <span data-ttu-id="c0c87-171">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-171">Type</span></span> | <span data-ttu-id="c0c87-172">说明</span><span class="sxs-lookup"><span data-stu-id="c0c87-172">Description</span></span> | <span data-ttu-id="c0c87-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="c0c87-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="c0c87-174">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-174">String</span></span> | <span data-ttu-id="c0c87-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="c0c87-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c0c87-176">1.7</span><span class="sxs-lookup"><span data-stu-id="c0c87-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="c0c87-177">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-177">String</span></span> | <span data-ttu-id="c0c87-178">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="c0c87-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="c0c87-179">1.8</span><span class="sxs-lookup"><span data-stu-id="c0c87-179">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="c0c87-180">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-180">String</span></span> | <span data-ttu-id="c0c87-181">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="c0c87-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="c0c87-182">1.8</span><span class="sxs-lookup"><span data-stu-id="c0c87-182">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="c0c87-183">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-183">String</span></span> | <span data-ttu-id="c0c87-184">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="c0c87-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="c0c87-185">1.5</span><span class="sxs-lookup"><span data-stu-id="c0c87-185">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c0c87-186">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-186">String</span></span> | <span data-ttu-id="c0c87-187">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="c0c87-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c0c87-188">1.7</span><span class="sxs-lookup"><span data-stu-id="c0c87-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c0c87-189">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-189">String</span></span> | <span data-ttu-id="c0c87-190">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="c0c87-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c0c87-191">1.7</span><span class="sxs-lookup"><span data-stu-id="c0c87-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c0c87-192">Requirements</span><span class="sxs-lookup"><span data-stu-id="c0c87-192">Requirements</span></span>

|<span data-ttu-id="c0c87-193">要求</span><span class="sxs-lookup"><span data-stu-id="c0c87-193">Requirement</span></span>| <span data-ttu-id="c0c87-194">值</span><span class="sxs-lookup"><span data-stu-id="c0c87-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c87-195">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c87-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c87-196">1.5</span><span class="sxs-lookup"><span data-stu-id="c0c87-196">1.5</span></span> |
|[<span data-ttu-id="c0c87-197">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c87-197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c87-198">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c87-198">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="c0c87-199">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="c0c87-199">SourceProperty: String</span></span>

<span data-ttu-id="c0c87-200">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="c0c87-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c0c87-201">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-201">Type</span></span>

*   <span data-ttu-id="c0c87-202">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c0c87-203">属性：</span><span class="sxs-lookup"><span data-stu-id="c0c87-203">Properties:</span></span>

|<span data-ttu-id="c0c87-204">名称</span><span class="sxs-lookup"><span data-stu-id="c0c87-204">Name</span></span>| <span data-ttu-id="c0c87-205">类型</span><span class="sxs-lookup"><span data-stu-id="c0c87-205">Type</span></span>| <span data-ttu-id="c0c87-206">说明</span><span class="sxs-lookup"><span data-stu-id="c0c87-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c0c87-207">字符串</span><span class="sxs-lookup"><span data-stu-id="c0c87-207">String</span></span>|<span data-ttu-id="c0c87-208">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="c0c87-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c0c87-209">String</span><span class="sxs-lookup"><span data-stu-id="c0c87-209">String</span></span>|<span data-ttu-id="c0c87-210">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="c0c87-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c0c87-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="c0c87-211">Requirements</span></span>

|<span data-ttu-id="c0c87-212">要求</span><span class="sxs-lookup"><span data-stu-id="c0c87-212">Requirement</span></span>| <span data-ttu-id="c0c87-213">值</span><span class="sxs-lookup"><span data-stu-id="c0c87-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0c87-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0c87-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0c87-215">1.0</span><span class="sxs-lookup"><span data-stu-id="c0c87-215">1.0</span></span>|
|[<span data-ttu-id="c0c87-216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0c87-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0c87-217">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0c87-217">Compose or Read</span></span>|
