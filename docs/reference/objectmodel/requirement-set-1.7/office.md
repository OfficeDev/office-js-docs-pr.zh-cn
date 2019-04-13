---
title: Office 命名空间-要求集1。7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 523df189b28fc568ac32e8d17d4a226b52cbd23c
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838450"
---
# <a name="office"></a><span data-ttu-id="d83c9-102">Office</span><span class="sxs-lookup"><span data-stu-id="d83c9-102">Office</span></span>

<span data-ttu-id="d83c9-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="d83c9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d83c9-105">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-105">Requirements</span></span>

|<span data-ttu-id="d83c9-106">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-106">Requirement</span></span>| <span data-ttu-id="d83c9-107">值</span><span class="sxs-lookup"><span data-stu-id="d83c9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d83c9-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d83c9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d83c9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d83c9-109">1.0</span></span>|
|[<span data-ttu-id="d83c9-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d83c9-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d83c9-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d83c9-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d83c9-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="d83c9-112">Members and methods</span></span>

| <span data-ttu-id="d83c9-113">成员</span><span class="sxs-lookup"><span data-stu-id="d83c9-113">Member</span></span> | <span data-ttu-id="d83c9-114">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d83c9-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d83c9-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d83c9-116">Member</span><span class="sxs-lookup"><span data-stu-id="d83c9-116">Member</span></span> |
| [<span data-ttu-id="d83c9-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d83c9-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d83c9-118">Member</span><span class="sxs-lookup"><span data-stu-id="d83c9-118">Member</span></span> |
| [<span data-ttu-id="d83c9-119">EventType</span><span class="sxs-lookup"><span data-stu-id="d83c9-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d83c9-120">Member</span><span class="sxs-lookup"><span data-stu-id="d83c9-120">Member</span></span> |
| [<span data-ttu-id="d83c9-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d83c9-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d83c9-122">成员</span><span class="sxs-lookup"><span data-stu-id="d83c9-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d83c9-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="d83c9-123">Namespaces</span></span>

<span data-ttu-id="d83c9-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="d83c9-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d83c9-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="d83c9-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d83c9-126">成员</span><span class="sxs-lookup"><span data-stu-id="d83c9-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d83c9-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d83c9-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="d83c9-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="d83c9-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d83c9-129">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-129">Type</span></span>

*   <span data-ttu-id="d83c9-130">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d83c9-131">属性：</span><span class="sxs-lookup"><span data-stu-id="d83c9-131">Properties:</span></span>

|<span data-ttu-id="d83c9-132">名称</span><span class="sxs-lookup"><span data-stu-id="d83c9-132">Name</span></span>| <span data-ttu-id="d83c9-133">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-133">Type</span></span>| <span data-ttu-id="d83c9-134">说明</span><span class="sxs-lookup"><span data-stu-id="d83c9-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d83c9-135">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-135">String</span></span>|<span data-ttu-id="d83c9-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="d83c9-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d83c9-137">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-137">String</span></span>|<span data-ttu-id="d83c9-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="d83c9-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d83c9-139">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-139">Requirements</span></span>

|<span data-ttu-id="d83c9-140">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-140">Requirement</span></span>| <span data-ttu-id="d83c9-141">值</span><span class="sxs-lookup"><span data-stu-id="d83c9-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="d83c9-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d83c9-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d83c9-143">1.0</span><span class="sxs-lookup"><span data-stu-id="d83c9-143">1.0</span></span>|
|[<span data-ttu-id="d83c9-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d83c9-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d83c9-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d83c9-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="d83c9-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d83c9-146">CoercionType :String</span></span>

<span data-ttu-id="d83c9-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="d83c9-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d83c9-148">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-148">Type</span></span>

*   <span data-ttu-id="d83c9-149">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d83c9-150">属性：</span><span class="sxs-lookup"><span data-stu-id="d83c9-150">Properties:</span></span>

|<span data-ttu-id="d83c9-151">名称</span><span class="sxs-lookup"><span data-stu-id="d83c9-151">Name</span></span>| <span data-ttu-id="d83c9-152">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-152">Type</span></span>| <span data-ttu-id="d83c9-153">说明</span><span class="sxs-lookup"><span data-stu-id="d83c9-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d83c9-154">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-154">String</span></span>|<span data-ttu-id="d83c9-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d83c9-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d83c9-156">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-156">String</span></span>|<span data-ttu-id="d83c9-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d83c9-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d83c9-158">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-158">Requirements</span></span>

|<span data-ttu-id="d83c9-159">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-159">Requirement</span></span>| <span data-ttu-id="d83c9-160">值</span><span class="sxs-lookup"><span data-stu-id="d83c9-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="d83c9-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d83c9-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d83c9-162">1.0</span><span class="sxs-lookup"><span data-stu-id="d83c9-162">1.0</span></span>|
|[<span data-ttu-id="d83c9-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d83c9-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d83c9-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d83c9-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="d83c9-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="d83c9-165">EventType :String</span></span>

<span data-ttu-id="d83c9-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="d83c9-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d83c9-167">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-167">Type</span></span>

*   <span data-ttu-id="d83c9-168">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d83c9-169">属性：</span><span class="sxs-lookup"><span data-stu-id="d83c9-169">Properties:</span></span>

| <span data-ttu-id="d83c9-170">名称</span><span class="sxs-lookup"><span data-stu-id="d83c9-170">Name</span></span> | <span data-ttu-id="d83c9-171">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-171">Type</span></span> | <span data-ttu-id="d83c9-172">说明</span><span class="sxs-lookup"><span data-stu-id="d83c9-172">Description</span></span> | <span data-ttu-id="d83c9-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="d83c9-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="d83c9-174">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-174">String</span></span> | <span data-ttu-id="d83c9-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="d83c9-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="d83c9-176">1.7</span><span class="sxs-lookup"><span data-stu-id="d83c9-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="d83c9-177">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-177">String</span></span> | <span data-ttu-id="d83c9-178">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="d83c9-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="d83c9-179">1.5</span><span class="sxs-lookup"><span data-stu-id="d83c9-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="d83c9-180">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-180">String</span></span> | <span data-ttu-id="d83c9-181">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="d83c9-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="d83c9-182">1.7</span><span class="sxs-lookup"><span data-stu-id="d83c9-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="d83c9-183">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-183">String</span></span> | <span data-ttu-id="d83c9-184">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="d83c9-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="d83c9-185">1.7</span><span class="sxs-lookup"><span data-stu-id="d83c9-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d83c9-186">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-186">Requirements</span></span>

|<span data-ttu-id="d83c9-187">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-187">Requirement</span></span>| <span data-ttu-id="d83c9-188">值</span><span class="sxs-lookup"><span data-stu-id="d83c9-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="d83c9-189">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d83c9-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d83c9-190">1.5</span><span class="sxs-lookup"><span data-stu-id="d83c9-190">1.5</span></span> |
|[<span data-ttu-id="d83c9-191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d83c9-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d83c9-192">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d83c9-192">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="d83c9-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d83c9-193">SourceProperty :String</span></span>

<span data-ttu-id="d83c9-194">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="d83c9-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d83c9-195">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-195">Type</span></span>

*   <span data-ttu-id="d83c9-196">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d83c9-197">属性：</span><span class="sxs-lookup"><span data-stu-id="d83c9-197">Properties:</span></span>

|<span data-ttu-id="d83c9-198">名称</span><span class="sxs-lookup"><span data-stu-id="d83c9-198">Name</span></span>| <span data-ttu-id="d83c9-199">类型</span><span class="sxs-lookup"><span data-stu-id="d83c9-199">Type</span></span>| <span data-ttu-id="d83c9-200">说明</span><span class="sxs-lookup"><span data-stu-id="d83c9-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d83c9-201">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-201">String</span></span>|<span data-ttu-id="d83c9-202">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="d83c9-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d83c9-203">String</span><span class="sxs-lookup"><span data-stu-id="d83c9-203">String</span></span>|<span data-ttu-id="d83c9-204">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="d83c9-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d83c9-205">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-205">Requirements</span></span>

|<span data-ttu-id="d83c9-206">要求</span><span class="sxs-lookup"><span data-stu-id="d83c9-206">Requirement</span></span>| <span data-ttu-id="d83c9-207">值</span><span class="sxs-lookup"><span data-stu-id="d83c9-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="d83c9-208">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d83c9-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d83c9-209">1.0</span><span class="sxs-lookup"><span data-stu-id="d83c9-209">1.0</span></span>|
|[<span data-ttu-id="d83c9-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d83c9-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d83c9-211">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d83c9-211">Compose or Read</span></span>|
