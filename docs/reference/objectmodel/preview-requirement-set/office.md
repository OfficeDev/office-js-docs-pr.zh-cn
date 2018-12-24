---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: a276af19ebd1816ad6bd59af5a75c39f13aa0b3c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432893"
---
# <a name="office"></a><span data-ttu-id="c8425-102">Office</span><span class="sxs-lookup"><span data-stu-id="c8425-102">Office</span></span>

<span data-ttu-id="c8425-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="c8425-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c8425-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="c8425-105">Requirements</span></span>

|<span data-ttu-id="c8425-106">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-106">Requirement</span></span>| <span data-ttu-id="c8425-107">值</span><span class="sxs-lookup"><span data-stu-id="c8425-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c8425-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c8425-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c8425-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c8425-109">1.0</span></span>|
|[<span data-ttu-id="c8425-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c8425-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c8425-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c8425-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c8425-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c8425-112">Members and methods</span></span>

| <span data-ttu-id="c8425-113">成员</span><span class="sxs-lookup"><span data-stu-id="c8425-113">Member</span></span> | <span data-ttu-id="c8425-114">类型</span><span class="sxs-lookup"><span data-stu-id="c8425-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c8425-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c8425-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c8425-116">成员</span><span class="sxs-lookup"><span data-stu-id="c8425-116">Member</span></span> |
| [<span data-ttu-id="c8425-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c8425-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c8425-118">成员</span><span class="sxs-lookup"><span data-stu-id="c8425-118">Member</span></span> |
| [<span data-ttu-id="c8425-119">EventType</span><span class="sxs-lookup"><span data-stu-id="c8425-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c8425-120">成员</span><span class="sxs-lookup"><span data-stu-id="c8425-120">Member</span></span> |
| [<span data-ttu-id="c8425-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c8425-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c8425-122">成员</span><span class="sxs-lookup"><span data-stu-id="c8425-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c8425-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="c8425-123">Namespaces</span></span>

<span data-ttu-id="c8425-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="c8425-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c8425-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="c8425-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c8425-126">成员</span><span class="sxs-lookup"><span data-stu-id="c8425-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c8425-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c8425-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="c8425-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="c8425-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c8425-129">类型：</span><span class="sxs-lookup"><span data-stu-id="c8425-129">Type:</span></span>

*   <span data-ttu-id="c8425-130">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c8425-131">属性：</span><span class="sxs-lookup"><span data-stu-id="c8425-131">Properties:</span></span>

|<span data-ttu-id="c8425-132">名称</span><span class="sxs-lookup"><span data-stu-id="c8425-132">Name</span></span>| <span data-ttu-id="c8425-133">类型</span><span class="sxs-lookup"><span data-stu-id="c8425-133">Type</span></span>| <span data-ttu-id="c8425-134">描述</span><span class="sxs-lookup"><span data-stu-id="c8425-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c8425-135">String</span><span class="sxs-lookup"><span data-stu-id="c8425-135">String</span></span>|<span data-ttu-id="c8425-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="c8425-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c8425-137">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-137">String</span></span>|<span data-ttu-id="c8425-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="c8425-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c8425-139">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-139">Requirements</span></span>

|<span data-ttu-id="c8425-140">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-140">Requirement</span></span>| <span data-ttu-id="c8425-141">值</span><span class="sxs-lookup"><span data-stu-id="c8425-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="c8425-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c8425-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c8425-143">1.0</span><span class="sxs-lookup"><span data-stu-id="c8425-143">1.0</span></span>|
|[<span data-ttu-id="c8425-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c8425-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c8425-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c8425-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="c8425-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c8425-146">CoercionType :String</span></span>

<span data-ttu-id="c8425-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="c8425-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c8425-148">类型：</span><span class="sxs-lookup"><span data-stu-id="c8425-148">Type:</span></span>

*   <span data-ttu-id="c8425-149">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c8425-150">属性：</span><span class="sxs-lookup"><span data-stu-id="c8425-150">Properties:</span></span>

|<span data-ttu-id="c8425-151">名称</span><span class="sxs-lookup"><span data-stu-id="c8425-151">Name</span></span>| <span data-ttu-id="c8425-152">类型</span><span class="sxs-lookup"><span data-stu-id="c8425-152">Type</span></span>| <span data-ttu-id="c8425-153">描述</span><span class="sxs-lookup"><span data-stu-id="c8425-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c8425-154">String</span><span class="sxs-lookup"><span data-stu-id="c8425-154">String</span></span>|<span data-ttu-id="c8425-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="c8425-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c8425-156">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-156">String</span></span>|<span data-ttu-id="c8425-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="c8425-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c8425-158">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-158">Requirements</span></span>

|<span data-ttu-id="c8425-159">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-159">Requirement</span></span>| <span data-ttu-id="c8425-160">值</span><span class="sxs-lookup"><span data-stu-id="c8425-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="c8425-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c8425-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c8425-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c8425-162">1.0</span></span>|
|[<span data-ttu-id="c8425-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c8425-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c8425-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c8425-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="c8425-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="c8425-165">EventType :String</span></span>

<span data-ttu-id="c8425-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="c8425-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c8425-167">类型：</span><span class="sxs-lookup"><span data-stu-id="c8425-167">Type:</span></span>

*   <span data-ttu-id="c8425-168">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c8425-169">属性：</span><span class="sxs-lookup"><span data-stu-id="c8425-169">Properties:</span></span>

| <span data-ttu-id="c8425-170">名称</span><span class="sxs-lookup"><span data-stu-id="c8425-170">Name</span></span> | <span data-ttu-id="c8425-171">类型</span><span class="sxs-lookup"><span data-stu-id="c8425-171">Type</span></span> | <span data-ttu-id="c8425-172">描述</span><span class="sxs-lookup"><span data-stu-id="c8425-172">Description</span></span> | <span data-ttu-id="c8425-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="c8425-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="c8425-174">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-174">String</span></span> | <span data-ttu-id="c8425-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="c8425-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c8425-176">1.7</span><span class="sxs-lookup"><span data-stu-id="c8425-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="c8425-177">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-177">String</span></span> | <span data-ttu-id="c8425-178">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="c8425-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="c8425-179">预览</span><span class="sxs-lookup"><span data-stu-id="c8425-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="c8425-180">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-180">String</span></span> | <span data-ttu-id="c8425-181">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="c8425-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="c8425-182">1.5</span><span class="sxs-lookup"><span data-stu-id="c8425-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="c8425-183">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-183">String</span></span> | <span data-ttu-id="c8425-184">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="c8425-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="c8425-185">预览</span><span class="sxs-lookup"><span data-stu-id="c8425-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c8425-186">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-186">String</span></span> | <span data-ttu-id="c8425-187">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="c8425-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c8425-188">1.7</span><span class="sxs-lookup"><span data-stu-id="c8425-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c8425-189">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-189">String</span></span> | <span data-ttu-id="c8425-190">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="c8425-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c8425-191">1.7</span><span class="sxs-lookup"><span data-stu-id="c8425-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c8425-192">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-192">Requirements</span></span>

|<span data-ttu-id="c8425-193">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-193">Requirement</span></span>| <span data-ttu-id="c8425-194">值</span><span class="sxs-lookup"><span data-stu-id="c8425-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="c8425-195">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c8425-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c8425-196">1.5</span><span class="sxs-lookup"><span data-stu-id="c8425-196">1.5</span></span> |
|[<span data-ttu-id="c8425-197">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c8425-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c8425-198">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c8425-198">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="c8425-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c8425-199">SourceProperty :String</span></span>

<span data-ttu-id="c8425-200">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="c8425-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c8425-201">类型：</span><span class="sxs-lookup"><span data-stu-id="c8425-201">Type:</span></span>

*   <span data-ttu-id="c8425-202">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c8425-203">属性：</span><span class="sxs-lookup"><span data-stu-id="c8425-203">Properties:</span></span>

|<span data-ttu-id="c8425-204">名称</span><span class="sxs-lookup"><span data-stu-id="c8425-204">Name</span></span>| <span data-ttu-id="c8425-205">类型</span><span class="sxs-lookup"><span data-stu-id="c8425-205">Type</span></span>| <span data-ttu-id="c8425-206">描述</span><span class="sxs-lookup"><span data-stu-id="c8425-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c8425-207">字符串</span><span class="sxs-lookup"><span data-stu-id="c8425-207">String</span></span>|<span data-ttu-id="c8425-208">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="c8425-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c8425-209">String</span><span class="sxs-lookup"><span data-stu-id="c8425-209">String</span></span>|<span data-ttu-id="c8425-210">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="c8425-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c8425-211">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-211">Requirements</span></span>

|<span data-ttu-id="c8425-212">要求</span><span class="sxs-lookup"><span data-stu-id="c8425-212">Requirement</span></span>| <span data-ttu-id="c8425-213">值</span><span class="sxs-lookup"><span data-stu-id="c8425-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="c8425-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c8425-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c8425-215">1.0</span><span class="sxs-lookup"><span data-stu-id="c8425-215">1.0</span></span>|
|[<span data-ttu-id="c8425-216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c8425-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c8425-217">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c8425-217">Compose or read</span></span>|