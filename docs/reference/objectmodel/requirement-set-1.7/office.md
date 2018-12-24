---
title: Office 命名空间 - 要求集 1.7
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 2bf1c31f4dc4156cb4f1d0eb3508193305c860e9
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432800"
---
# <a name="office"></a><span data-ttu-id="d9df3-102">Office</span><span class="sxs-lookup"><span data-stu-id="d9df3-102">Office</span></span>

<span data-ttu-id="d9df3-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="d9df3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9df3-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="d9df3-105">Requirements</span></span>

|<span data-ttu-id="d9df3-106">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-106">Requirement</span></span>| <span data-ttu-id="d9df3-107">值</span><span class="sxs-lookup"><span data-stu-id="d9df3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9df3-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d9df3-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9df3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d9df3-109">1.0</span></span>|
|[<span data-ttu-id="d9df3-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d9df3-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9df3-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d9df3-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d9df3-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="d9df3-112">Members and methods</span></span>

| <span data-ttu-id="d9df3-113">成员</span><span class="sxs-lookup"><span data-stu-id="d9df3-113">Member</span></span> | <span data-ttu-id="d9df3-114">类型</span><span class="sxs-lookup"><span data-stu-id="d9df3-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d9df3-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d9df3-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d9df3-116">成员</span><span class="sxs-lookup"><span data-stu-id="d9df3-116">Member</span></span> |
| [<span data-ttu-id="d9df3-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d9df3-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d9df3-118">成员</span><span class="sxs-lookup"><span data-stu-id="d9df3-118">Member</span></span> |
| [<span data-ttu-id="d9df3-119">EventType</span><span class="sxs-lookup"><span data-stu-id="d9df3-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d9df3-120">成员</span><span class="sxs-lookup"><span data-stu-id="d9df3-120">Member</span></span> |
| [<span data-ttu-id="d9df3-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d9df3-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d9df3-122">成员</span><span class="sxs-lookup"><span data-stu-id="d9df3-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d9df3-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="d9df3-123">Namespaces</span></span>

<span data-ttu-id="d9df3-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="d9df3-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d9df3-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="d9df3-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d9df3-126">成员</span><span class="sxs-lookup"><span data-stu-id="d9df3-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d9df3-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d9df3-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="d9df3-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="d9df3-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d9df3-129">类型：</span><span class="sxs-lookup"><span data-stu-id="d9df3-129">Type:</span></span>

*   <span data-ttu-id="d9df3-130">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d9df3-131">属性：</span><span class="sxs-lookup"><span data-stu-id="d9df3-131">Properties:</span></span>

|<span data-ttu-id="d9df3-132">名称</span><span class="sxs-lookup"><span data-stu-id="d9df3-132">Name</span></span>| <span data-ttu-id="d9df3-133">类型</span><span class="sxs-lookup"><span data-stu-id="d9df3-133">Type</span></span>| <span data-ttu-id="d9df3-134">描述</span><span class="sxs-lookup"><span data-stu-id="d9df3-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d9df3-135">String</span><span class="sxs-lookup"><span data-stu-id="d9df3-135">String</span></span>|<span data-ttu-id="d9df3-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="d9df3-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d9df3-137">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-137">String</span></span>|<span data-ttu-id="d9df3-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="d9df3-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9df3-139">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-139">Requirements</span></span>

|<span data-ttu-id="d9df3-140">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-140">Requirement</span></span>| <span data-ttu-id="d9df3-141">值</span><span class="sxs-lookup"><span data-stu-id="d9df3-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9df3-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d9df3-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9df3-143">1.0</span><span class="sxs-lookup"><span data-stu-id="d9df3-143">1.0</span></span>|
|[<span data-ttu-id="d9df3-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d9df3-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9df3-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d9df3-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="d9df3-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d9df3-146">CoercionType :String</span></span>

<span data-ttu-id="d9df3-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="d9df3-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d9df3-148">类型：</span><span class="sxs-lookup"><span data-stu-id="d9df3-148">Type:</span></span>

*   <span data-ttu-id="d9df3-149">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d9df3-150">属性：</span><span class="sxs-lookup"><span data-stu-id="d9df3-150">Properties:</span></span>

|<span data-ttu-id="d9df3-151">名称</span><span class="sxs-lookup"><span data-stu-id="d9df3-151">Name</span></span>| <span data-ttu-id="d9df3-152">类型</span><span class="sxs-lookup"><span data-stu-id="d9df3-152">Type</span></span>| <span data-ttu-id="d9df3-153">描述</span><span class="sxs-lookup"><span data-stu-id="d9df3-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d9df3-154">String</span><span class="sxs-lookup"><span data-stu-id="d9df3-154">String</span></span>|<span data-ttu-id="d9df3-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d9df3-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d9df3-156">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-156">String</span></span>|<span data-ttu-id="d9df3-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d9df3-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9df3-158">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-158">Requirements</span></span>

|<span data-ttu-id="d9df3-159">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-159">Requirement</span></span>| <span data-ttu-id="d9df3-160">值</span><span class="sxs-lookup"><span data-stu-id="d9df3-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9df3-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d9df3-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9df3-162">1.0</span><span class="sxs-lookup"><span data-stu-id="d9df3-162">1.0</span></span>|
|[<span data-ttu-id="d9df3-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d9df3-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9df3-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d9df3-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="d9df3-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="d9df3-165">EventType :String</span></span>

<span data-ttu-id="d9df3-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="d9df3-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d9df3-167">类型：</span><span class="sxs-lookup"><span data-stu-id="d9df3-167">Type:</span></span>

*   <span data-ttu-id="d9df3-168">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d9df3-169">属性：</span><span class="sxs-lookup"><span data-stu-id="d9df3-169">Properties:</span></span>

| <span data-ttu-id="d9df3-170">名称</span><span class="sxs-lookup"><span data-stu-id="d9df3-170">Name</span></span> | <span data-ttu-id="d9df3-171">类型</span><span class="sxs-lookup"><span data-stu-id="d9df3-171">Type</span></span> | <span data-ttu-id="d9df3-172">描述</span><span class="sxs-lookup"><span data-stu-id="d9df3-172">Description</span></span> | <span data-ttu-id="d9df3-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="d9df3-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="d9df3-174">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-174">String</span></span> | <span data-ttu-id="d9df3-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="d9df3-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="d9df3-176">1.7</span><span class="sxs-lookup"><span data-stu-id="d9df3-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="d9df3-177">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-177">String</span></span> | <span data-ttu-id="d9df3-178">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="d9df3-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="d9df3-179">1.5</span><span class="sxs-lookup"><span data-stu-id="d9df3-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="d9df3-180">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-180">String</span></span> | <span data-ttu-id="d9df3-181">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="d9df3-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="d9df3-182">1.7</span><span class="sxs-lookup"><span data-stu-id="d9df3-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="d9df3-183">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-183">String</span></span> | <span data-ttu-id="d9df3-184">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="d9df3-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="d9df3-185">1.7</span><span class="sxs-lookup"><span data-stu-id="d9df3-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9df3-186">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-186">Requirements</span></span>

|<span data-ttu-id="d9df3-187">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-187">Requirement</span></span>| <span data-ttu-id="d9df3-188">值</span><span class="sxs-lookup"><span data-stu-id="d9df3-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9df3-189">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d9df3-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9df3-190">1.5</span><span class="sxs-lookup"><span data-stu-id="d9df3-190">1.5</span></span> |
|[<span data-ttu-id="d9df3-191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d9df3-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9df3-192">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d9df3-192">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="d9df3-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d9df3-193">SourceProperty :String</span></span>

<span data-ttu-id="d9df3-194">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="d9df3-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d9df3-195">类型：</span><span class="sxs-lookup"><span data-stu-id="d9df3-195">Type:</span></span>

*   <span data-ttu-id="d9df3-196">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d9df3-197">属性：</span><span class="sxs-lookup"><span data-stu-id="d9df3-197">Properties:</span></span>

|<span data-ttu-id="d9df3-198">名称</span><span class="sxs-lookup"><span data-stu-id="d9df3-198">Name</span></span>| <span data-ttu-id="d9df3-199">类型</span><span class="sxs-lookup"><span data-stu-id="d9df3-199">Type</span></span>| <span data-ttu-id="d9df3-200">描述</span><span class="sxs-lookup"><span data-stu-id="d9df3-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d9df3-201">字符串</span><span class="sxs-lookup"><span data-stu-id="d9df3-201">String</span></span>|<span data-ttu-id="d9df3-202">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="d9df3-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d9df3-203">String</span><span class="sxs-lookup"><span data-stu-id="d9df3-203">String</span></span>|<span data-ttu-id="d9df3-204">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="d9df3-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9df3-205">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-205">Requirements</span></span>

|<span data-ttu-id="d9df3-206">要求</span><span class="sxs-lookup"><span data-stu-id="d9df3-206">Requirement</span></span>| <span data-ttu-id="d9df3-207">值</span><span class="sxs-lookup"><span data-stu-id="d9df3-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9df3-208">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d9df3-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9df3-209">1.0</span><span class="sxs-lookup"><span data-stu-id="d9df3-209">1.0</span></span>|
|[<span data-ttu-id="d9df3-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d9df3-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9df3-211">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d9df3-211">Compose or read</span></span>|