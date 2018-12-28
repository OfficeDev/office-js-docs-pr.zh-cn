---
title: Office 命名空间 - 要求集 1.7
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 6afaca31dd941b9c6a4b23fa08018de51278cbbd
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457738"
---
# <a name="office"></a><span data-ttu-id="71bad-102">Office</span><span class="sxs-lookup"><span data-stu-id="71bad-102">Office</span></span>

<span data-ttu-id="71bad-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="71bad-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="71bad-105">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-105">Requirements</span></span>

|<span data-ttu-id="71bad-106">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-106">Requirement</span></span>| <span data-ttu-id="71bad-107">值</span><span class="sxs-lookup"><span data-stu-id="71bad-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="71bad-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71bad-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71bad-109">1.0</span><span class="sxs-lookup"><span data-stu-id="71bad-109">1.0</span></span>|
|[<span data-ttu-id="71bad-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71bad-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71bad-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71bad-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="71bad-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="71bad-112">Members and methods</span></span>

| <span data-ttu-id="71bad-113">成员</span><span class="sxs-lookup"><span data-stu-id="71bad-113">Member</span></span> | <span data-ttu-id="71bad-114">类型</span><span class="sxs-lookup"><span data-stu-id="71bad-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="71bad-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="71bad-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="71bad-116">成员</span><span class="sxs-lookup"><span data-stu-id="71bad-116">Member</span></span> |
| [<span data-ttu-id="71bad-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="71bad-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="71bad-118">成员</span><span class="sxs-lookup"><span data-stu-id="71bad-118">Member</span></span> |
| [<span data-ttu-id="71bad-119">EventType</span><span class="sxs-lookup"><span data-stu-id="71bad-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="71bad-120">成员</span><span class="sxs-lookup"><span data-stu-id="71bad-120">Member</span></span> |
| [<span data-ttu-id="71bad-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="71bad-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="71bad-122">成员</span><span class="sxs-lookup"><span data-stu-id="71bad-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="71bad-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="71bad-123">Namespaces</span></span>

<span data-ttu-id="71bad-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="71bad-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="71bad-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="71bad-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="71bad-126">成员</span><span class="sxs-lookup"><span data-stu-id="71bad-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="71bad-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="71bad-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="71bad-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="71bad-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="71bad-129">类型：</span><span class="sxs-lookup"><span data-stu-id="71bad-129">Type:</span></span>

*   <span data-ttu-id="71bad-130">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71bad-131">属性：</span><span class="sxs-lookup"><span data-stu-id="71bad-131">Properties:</span></span>

|<span data-ttu-id="71bad-132">名称</span><span class="sxs-lookup"><span data-stu-id="71bad-132">Name</span></span>| <span data-ttu-id="71bad-133">类型</span><span class="sxs-lookup"><span data-stu-id="71bad-133">Type</span></span>| <span data-ttu-id="71bad-134">描述</span><span class="sxs-lookup"><span data-stu-id="71bad-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="71bad-135">String</span><span class="sxs-lookup"><span data-stu-id="71bad-135">String</span></span>|<span data-ttu-id="71bad-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="71bad-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="71bad-137">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-137">String</span></span>|<span data-ttu-id="71bad-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="71bad-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71bad-139">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-139">Requirements</span></span>

|<span data-ttu-id="71bad-140">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-140">Requirement</span></span>| <span data-ttu-id="71bad-141">值</span><span class="sxs-lookup"><span data-stu-id="71bad-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="71bad-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71bad-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71bad-143">1.0</span><span class="sxs-lookup"><span data-stu-id="71bad-143">1.0</span></span>|
|[<span data-ttu-id="71bad-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71bad-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71bad-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71bad-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="71bad-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="71bad-146">CoercionType :String</span></span>

<span data-ttu-id="71bad-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="71bad-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="71bad-148">类型：</span><span class="sxs-lookup"><span data-stu-id="71bad-148">Type:</span></span>

*   <span data-ttu-id="71bad-149">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71bad-150">属性：</span><span class="sxs-lookup"><span data-stu-id="71bad-150">Properties:</span></span>

|<span data-ttu-id="71bad-151">名称</span><span class="sxs-lookup"><span data-stu-id="71bad-151">Name</span></span>| <span data-ttu-id="71bad-152">类型</span><span class="sxs-lookup"><span data-stu-id="71bad-152">Type</span></span>| <span data-ttu-id="71bad-153">描述</span><span class="sxs-lookup"><span data-stu-id="71bad-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="71bad-154">String</span><span class="sxs-lookup"><span data-stu-id="71bad-154">String</span></span>|<span data-ttu-id="71bad-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="71bad-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="71bad-156">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-156">String</span></span>|<span data-ttu-id="71bad-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="71bad-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71bad-158">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-158">Requirements</span></span>

|<span data-ttu-id="71bad-159">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-159">Requirement</span></span>| <span data-ttu-id="71bad-160">值</span><span class="sxs-lookup"><span data-stu-id="71bad-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="71bad-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71bad-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71bad-162">1.0</span><span class="sxs-lookup"><span data-stu-id="71bad-162">1.0</span></span>|
|[<span data-ttu-id="71bad-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71bad-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71bad-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71bad-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="71bad-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="71bad-165">EventType :String</span></span>

<span data-ttu-id="71bad-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="71bad-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="71bad-167">类型：</span><span class="sxs-lookup"><span data-stu-id="71bad-167">Type:</span></span>

*   <span data-ttu-id="71bad-168">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71bad-169">属性：</span><span class="sxs-lookup"><span data-stu-id="71bad-169">Properties:</span></span>

| <span data-ttu-id="71bad-170">名称</span><span class="sxs-lookup"><span data-stu-id="71bad-170">Name</span></span> | <span data-ttu-id="71bad-171">类型</span><span class="sxs-lookup"><span data-stu-id="71bad-171">Type</span></span> | <span data-ttu-id="71bad-172">描述</span><span class="sxs-lookup"><span data-stu-id="71bad-172">Description</span></span> | <span data-ttu-id="71bad-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="71bad-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="71bad-174">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-174">String</span></span> | <span data-ttu-id="71bad-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="71bad-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="71bad-176">1.7</span><span class="sxs-lookup"><span data-stu-id="71bad-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="71bad-177">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-177">String</span></span> | <span data-ttu-id="71bad-178">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="71bad-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="71bad-179">1.5</span><span class="sxs-lookup"><span data-stu-id="71bad-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="71bad-180">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-180">String</span></span> | <span data-ttu-id="71bad-181">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="71bad-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="71bad-182">1.7</span><span class="sxs-lookup"><span data-stu-id="71bad-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="71bad-183">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-183">String</span></span> | <span data-ttu-id="71bad-184">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="71bad-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="71bad-185">1.7</span><span class="sxs-lookup"><span data-stu-id="71bad-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="71bad-186">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-186">Requirements</span></span>

|<span data-ttu-id="71bad-187">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-187">Requirement</span></span>| <span data-ttu-id="71bad-188">值</span><span class="sxs-lookup"><span data-stu-id="71bad-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="71bad-189">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71bad-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71bad-190">1.5</span><span class="sxs-lookup"><span data-stu-id="71bad-190">1.5</span></span> |
|[<span data-ttu-id="71bad-191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71bad-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71bad-192">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71bad-192">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="71bad-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="71bad-193">SourceProperty :String</span></span>

<span data-ttu-id="71bad-194">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="71bad-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="71bad-195">类型：</span><span class="sxs-lookup"><span data-stu-id="71bad-195">Type:</span></span>

*   <span data-ttu-id="71bad-196">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71bad-197">属性：</span><span class="sxs-lookup"><span data-stu-id="71bad-197">Properties:</span></span>

|<span data-ttu-id="71bad-198">名称</span><span class="sxs-lookup"><span data-stu-id="71bad-198">Name</span></span>| <span data-ttu-id="71bad-199">类型</span><span class="sxs-lookup"><span data-stu-id="71bad-199">Type</span></span>| <span data-ttu-id="71bad-200">描述</span><span class="sxs-lookup"><span data-stu-id="71bad-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="71bad-201">字符串</span><span class="sxs-lookup"><span data-stu-id="71bad-201">String</span></span>|<span data-ttu-id="71bad-202">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="71bad-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="71bad-203">String</span><span class="sxs-lookup"><span data-stu-id="71bad-203">String</span></span>|<span data-ttu-id="71bad-204">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="71bad-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71bad-205">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-205">Requirements</span></span>

|<span data-ttu-id="71bad-206">要求</span><span class="sxs-lookup"><span data-stu-id="71bad-206">Requirement</span></span>| <span data-ttu-id="71bad-207">值</span><span class="sxs-lookup"><span data-stu-id="71bad-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="71bad-208">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71bad-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71bad-209">1.0</span><span class="sxs-lookup"><span data-stu-id="71bad-209">1.0</span></span>|
|[<span data-ttu-id="71bad-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71bad-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71bad-211">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71bad-211">Compose or read</span></span>|