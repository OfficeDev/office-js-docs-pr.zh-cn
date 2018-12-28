---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: f4a4f0d7a4ce0de433d4e70b6a4675b5f63f26f0
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457927"
---
# <a name="office"></a><span data-ttu-id="58264-102">Office</span><span class="sxs-lookup"><span data-stu-id="58264-102">Office</span></span>

<span data-ttu-id="58264-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="58264-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="58264-105">要求</span><span class="sxs-lookup"><span data-stu-id="58264-105">Requirements</span></span>

|<span data-ttu-id="58264-106">要求</span><span class="sxs-lookup"><span data-stu-id="58264-106">Requirement</span></span>| <span data-ttu-id="58264-107">值</span><span class="sxs-lookup"><span data-stu-id="58264-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="58264-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="58264-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58264-109">1.0</span><span class="sxs-lookup"><span data-stu-id="58264-109">1.0</span></span>|
|[<span data-ttu-id="58264-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="58264-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="58264-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="58264-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="58264-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="58264-112">Members and methods</span></span>

| <span data-ttu-id="58264-113">成员</span><span class="sxs-lookup"><span data-stu-id="58264-113">Member</span></span> | <span data-ttu-id="58264-114">类型</span><span class="sxs-lookup"><span data-stu-id="58264-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="58264-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="58264-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="58264-116">成员</span><span class="sxs-lookup"><span data-stu-id="58264-116">Member</span></span> |
| [<span data-ttu-id="58264-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="58264-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="58264-118">成员</span><span class="sxs-lookup"><span data-stu-id="58264-118">Member</span></span> |
| [<span data-ttu-id="58264-119">EventType</span><span class="sxs-lookup"><span data-stu-id="58264-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="58264-120">成员</span><span class="sxs-lookup"><span data-stu-id="58264-120">Member</span></span> |
| [<span data-ttu-id="58264-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="58264-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="58264-122">成员</span><span class="sxs-lookup"><span data-stu-id="58264-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="58264-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="58264-123">Namespaces</span></span>

<span data-ttu-id="58264-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="58264-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="58264-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="58264-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="58264-126">成员</span><span class="sxs-lookup"><span data-stu-id="58264-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="58264-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="58264-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="58264-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="58264-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="58264-129">类型：</span><span class="sxs-lookup"><span data-stu-id="58264-129">Type:</span></span>

*   <span data-ttu-id="58264-130">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="58264-131">属性：</span><span class="sxs-lookup"><span data-stu-id="58264-131">Properties:</span></span>

|<span data-ttu-id="58264-132">名称</span><span class="sxs-lookup"><span data-stu-id="58264-132">Name</span></span>| <span data-ttu-id="58264-133">类型</span><span class="sxs-lookup"><span data-stu-id="58264-133">Type</span></span>| <span data-ttu-id="58264-134">描述</span><span class="sxs-lookup"><span data-stu-id="58264-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="58264-135">String</span><span class="sxs-lookup"><span data-stu-id="58264-135">String</span></span>|<span data-ttu-id="58264-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="58264-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="58264-137">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-137">String</span></span>|<span data-ttu-id="58264-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="58264-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58264-139">要求</span><span class="sxs-lookup"><span data-stu-id="58264-139">Requirements</span></span>

|<span data-ttu-id="58264-140">要求</span><span class="sxs-lookup"><span data-stu-id="58264-140">Requirement</span></span>| <span data-ttu-id="58264-141">值</span><span class="sxs-lookup"><span data-stu-id="58264-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="58264-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="58264-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58264-143">1.0</span><span class="sxs-lookup"><span data-stu-id="58264-143">1.0</span></span>|
|[<span data-ttu-id="58264-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="58264-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="58264-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="58264-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="58264-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="58264-146">CoercionType :String</span></span>

<span data-ttu-id="58264-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="58264-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="58264-148">类型：</span><span class="sxs-lookup"><span data-stu-id="58264-148">Type:</span></span>

*   <span data-ttu-id="58264-149">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="58264-150">属性：</span><span class="sxs-lookup"><span data-stu-id="58264-150">Properties:</span></span>

|<span data-ttu-id="58264-151">名称</span><span class="sxs-lookup"><span data-stu-id="58264-151">Name</span></span>| <span data-ttu-id="58264-152">类型</span><span class="sxs-lookup"><span data-stu-id="58264-152">Type</span></span>| <span data-ttu-id="58264-153">描述</span><span class="sxs-lookup"><span data-stu-id="58264-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="58264-154">String</span><span class="sxs-lookup"><span data-stu-id="58264-154">String</span></span>|<span data-ttu-id="58264-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="58264-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="58264-156">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-156">String</span></span>|<span data-ttu-id="58264-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="58264-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58264-158">要求</span><span class="sxs-lookup"><span data-stu-id="58264-158">Requirements</span></span>

|<span data-ttu-id="58264-159">要求</span><span class="sxs-lookup"><span data-stu-id="58264-159">Requirement</span></span>| <span data-ttu-id="58264-160">值</span><span class="sxs-lookup"><span data-stu-id="58264-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="58264-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="58264-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58264-162">1.0</span><span class="sxs-lookup"><span data-stu-id="58264-162">1.0</span></span>|
|[<span data-ttu-id="58264-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="58264-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="58264-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="58264-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="58264-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="58264-165">EventType :String</span></span>

<span data-ttu-id="58264-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="58264-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="58264-167">类型：</span><span class="sxs-lookup"><span data-stu-id="58264-167">Type:</span></span>

*   <span data-ttu-id="58264-168">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="58264-169">属性：</span><span class="sxs-lookup"><span data-stu-id="58264-169">Properties:</span></span>

| <span data-ttu-id="58264-170">名称</span><span class="sxs-lookup"><span data-stu-id="58264-170">Name</span></span> | <span data-ttu-id="58264-171">类型</span><span class="sxs-lookup"><span data-stu-id="58264-171">Type</span></span> | <span data-ttu-id="58264-172">描述</span><span class="sxs-lookup"><span data-stu-id="58264-172">Description</span></span> | <span data-ttu-id="58264-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="58264-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="58264-174">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-174">String</span></span> | <span data-ttu-id="58264-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="58264-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="58264-176">1.7</span><span class="sxs-lookup"><span data-stu-id="58264-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="58264-177">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-177">String</span></span> | <span data-ttu-id="58264-178">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="58264-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="58264-179">预览</span><span class="sxs-lookup"><span data-stu-id="58264-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="58264-180">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-180">String</span></span> | <span data-ttu-id="58264-181">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="58264-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="58264-182">1.5</span><span class="sxs-lookup"><span data-stu-id="58264-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="58264-183">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-183">String</span></span> | <span data-ttu-id="58264-184">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="58264-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="58264-185">预览</span><span class="sxs-lookup"><span data-stu-id="58264-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="58264-186">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-186">String</span></span> | <span data-ttu-id="58264-187">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="58264-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="58264-188">1.7</span><span class="sxs-lookup"><span data-stu-id="58264-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="58264-189">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-189">String</span></span> | <span data-ttu-id="58264-190">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="58264-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="58264-191">1.7</span><span class="sxs-lookup"><span data-stu-id="58264-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="58264-192">要求</span><span class="sxs-lookup"><span data-stu-id="58264-192">Requirements</span></span>

|<span data-ttu-id="58264-193">要求</span><span class="sxs-lookup"><span data-stu-id="58264-193">Requirement</span></span>| <span data-ttu-id="58264-194">值</span><span class="sxs-lookup"><span data-stu-id="58264-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="58264-195">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="58264-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58264-196">1.5</span><span class="sxs-lookup"><span data-stu-id="58264-196">1.5</span></span> |
|[<span data-ttu-id="58264-197">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="58264-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="58264-198">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="58264-198">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="58264-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="58264-199">SourceProperty :String</span></span>

<span data-ttu-id="58264-200">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="58264-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="58264-201">类型：</span><span class="sxs-lookup"><span data-stu-id="58264-201">Type:</span></span>

*   <span data-ttu-id="58264-202">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="58264-203">属性：</span><span class="sxs-lookup"><span data-stu-id="58264-203">Properties:</span></span>

|<span data-ttu-id="58264-204">名称</span><span class="sxs-lookup"><span data-stu-id="58264-204">Name</span></span>| <span data-ttu-id="58264-205">类型</span><span class="sxs-lookup"><span data-stu-id="58264-205">Type</span></span>| <span data-ttu-id="58264-206">描述</span><span class="sxs-lookup"><span data-stu-id="58264-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="58264-207">字符串</span><span class="sxs-lookup"><span data-stu-id="58264-207">String</span></span>|<span data-ttu-id="58264-208">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="58264-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="58264-209">String</span><span class="sxs-lookup"><span data-stu-id="58264-209">String</span></span>|<span data-ttu-id="58264-210">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="58264-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58264-211">要求</span><span class="sxs-lookup"><span data-stu-id="58264-211">Requirements</span></span>

|<span data-ttu-id="58264-212">要求</span><span class="sxs-lookup"><span data-stu-id="58264-212">Requirement</span></span>| <span data-ttu-id="58264-213">值</span><span class="sxs-lookup"><span data-stu-id="58264-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="58264-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="58264-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58264-215">1.0</span><span class="sxs-lookup"><span data-stu-id="58264-215">1.0</span></span>|
|[<span data-ttu-id="58264-216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="58264-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="58264-217">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="58264-217">Compose or read</span></span>|