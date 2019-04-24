---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 7effc930d196aa009c3c779b702e082ae388fada
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451953"
---
# <a name="office"></a><span data-ttu-id="be02b-102">Office</span><span class="sxs-lookup"><span data-stu-id="be02b-102">Office</span></span>

<span data-ttu-id="be02b-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="be02b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="be02b-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="be02b-105">Requirements</span></span>

|<span data-ttu-id="be02b-106">要求</span><span class="sxs-lookup"><span data-stu-id="be02b-106">Requirement</span></span>| <span data-ttu-id="be02b-107">值</span><span class="sxs-lookup"><span data-stu-id="be02b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="be02b-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="be02b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be02b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="be02b-109">1.0</span></span>|
|[<span data-ttu-id="be02b-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be02b-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be02b-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be02b-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="be02b-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="be02b-112">Members and methods</span></span>

| <span data-ttu-id="be02b-113">成员</span><span class="sxs-lookup"><span data-stu-id="be02b-113">Member</span></span> | <span data-ttu-id="be02b-114">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="be02b-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="be02b-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="be02b-116">Member</span><span class="sxs-lookup"><span data-stu-id="be02b-116">Member</span></span> |
| [<span data-ttu-id="be02b-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="be02b-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="be02b-118">Member</span><span class="sxs-lookup"><span data-stu-id="be02b-118">Member</span></span> |
| [<span data-ttu-id="be02b-119">EventType</span><span class="sxs-lookup"><span data-stu-id="be02b-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="be02b-120">Member</span><span class="sxs-lookup"><span data-stu-id="be02b-120">Member</span></span> |
| [<span data-ttu-id="be02b-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="be02b-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="be02b-122">成员</span><span class="sxs-lookup"><span data-stu-id="be02b-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="be02b-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="be02b-123">Namespaces</span></span>

<span data-ttu-id="be02b-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="be02b-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="be02b-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="be02b-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="be02b-126">成员</span><span class="sxs-lookup"><span data-stu-id="be02b-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="be02b-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="be02b-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="be02b-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="be02b-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="be02b-129">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-129">Type</span></span>

*   <span data-ttu-id="be02b-130">String</span><span class="sxs-lookup"><span data-stu-id="be02b-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="be02b-131">属性：</span><span class="sxs-lookup"><span data-stu-id="be02b-131">Properties:</span></span>

|<span data-ttu-id="be02b-132">名称</span><span class="sxs-lookup"><span data-stu-id="be02b-132">Name</span></span>| <span data-ttu-id="be02b-133">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-133">Type</span></span>| <span data-ttu-id="be02b-134">描述</span><span class="sxs-lookup"><span data-stu-id="be02b-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="be02b-135">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-135">String</span></span>|<span data-ttu-id="be02b-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="be02b-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="be02b-137">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-137">String</span></span>|<span data-ttu-id="be02b-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="be02b-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be02b-139">Requirements</span><span class="sxs-lookup"><span data-stu-id="be02b-139">Requirements</span></span>

|<span data-ttu-id="be02b-140">要求</span><span class="sxs-lookup"><span data-stu-id="be02b-140">Requirement</span></span>| <span data-ttu-id="be02b-141">值</span><span class="sxs-lookup"><span data-stu-id="be02b-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="be02b-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="be02b-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be02b-143">1.0</span><span class="sxs-lookup"><span data-stu-id="be02b-143">1.0</span></span>|
|[<span data-ttu-id="be02b-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be02b-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be02b-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be02b-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="be02b-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="be02b-146">CoercionType :String</span></span>

<span data-ttu-id="be02b-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="be02b-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="be02b-148">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-148">Type</span></span>

*   <span data-ttu-id="be02b-149">String</span><span class="sxs-lookup"><span data-stu-id="be02b-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="be02b-150">属性：</span><span class="sxs-lookup"><span data-stu-id="be02b-150">Properties:</span></span>

|<span data-ttu-id="be02b-151">名称</span><span class="sxs-lookup"><span data-stu-id="be02b-151">Name</span></span>| <span data-ttu-id="be02b-152">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-152">Type</span></span>| <span data-ttu-id="be02b-153">描述</span><span class="sxs-lookup"><span data-stu-id="be02b-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="be02b-154">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-154">String</span></span>|<span data-ttu-id="be02b-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="be02b-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="be02b-156">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-156">String</span></span>|<span data-ttu-id="be02b-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="be02b-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be02b-158">Requirements</span><span class="sxs-lookup"><span data-stu-id="be02b-158">Requirements</span></span>

|<span data-ttu-id="be02b-159">要求</span><span class="sxs-lookup"><span data-stu-id="be02b-159">Requirement</span></span>| <span data-ttu-id="be02b-160">值</span><span class="sxs-lookup"><span data-stu-id="be02b-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="be02b-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="be02b-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be02b-162">1.0</span><span class="sxs-lookup"><span data-stu-id="be02b-162">1.0</span></span>|
|[<span data-ttu-id="be02b-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be02b-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be02b-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be02b-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="be02b-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="be02b-165">EventType :String</span></span>

<span data-ttu-id="be02b-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="be02b-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="be02b-167">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-167">Type</span></span>

*   <span data-ttu-id="be02b-168">String</span><span class="sxs-lookup"><span data-stu-id="be02b-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="be02b-169">属性：</span><span class="sxs-lookup"><span data-stu-id="be02b-169">Properties:</span></span>

| <span data-ttu-id="be02b-170">名称</span><span class="sxs-lookup"><span data-stu-id="be02b-170">Name</span></span> | <span data-ttu-id="be02b-171">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-171">Type</span></span> | <span data-ttu-id="be02b-172">描述</span><span class="sxs-lookup"><span data-stu-id="be02b-172">Description</span></span> | <span data-ttu-id="be02b-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="be02b-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="be02b-174">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-174">String</span></span> | <span data-ttu-id="be02b-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="be02b-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="be02b-176">1.7</span><span class="sxs-lookup"><span data-stu-id="be02b-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="be02b-177">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-177">String</span></span> | <span data-ttu-id="be02b-178">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="be02b-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="be02b-179">预览</span><span class="sxs-lookup"><span data-stu-id="be02b-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="be02b-180">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-180">String</span></span> | <span data-ttu-id="be02b-181">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="be02b-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="be02b-182">预览</span><span class="sxs-lookup"><span data-stu-id="be02b-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="be02b-183">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-183">String</span></span> | <span data-ttu-id="be02b-184">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="be02b-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="be02b-185">1.5</span><span class="sxs-lookup"><span data-stu-id="be02b-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="be02b-186">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-186">String</span></span> | <span data-ttu-id="be02b-187">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="be02b-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="be02b-188">预览</span><span class="sxs-lookup"><span data-stu-id="be02b-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="be02b-189">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-189">String</span></span> | <span data-ttu-id="be02b-190">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="be02b-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="be02b-191">1.7</span><span class="sxs-lookup"><span data-stu-id="be02b-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="be02b-192">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-192">String</span></span> | <span data-ttu-id="be02b-193">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="be02b-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="be02b-194">1.7</span><span class="sxs-lookup"><span data-stu-id="be02b-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="be02b-195">Requirements</span><span class="sxs-lookup"><span data-stu-id="be02b-195">Requirements</span></span>

|<span data-ttu-id="be02b-196">要求</span><span class="sxs-lookup"><span data-stu-id="be02b-196">Requirement</span></span>| <span data-ttu-id="be02b-197">值</span><span class="sxs-lookup"><span data-stu-id="be02b-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="be02b-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="be02b-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be02b-199">1.5</span><span class="sxs-lookup"><span data-stu-id="be02b-199">1.5</span></span> |
|[<span data-ttu-id="be02b-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be02b-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be02b-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be02b-201">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="be02b-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="be02b-202">SourceProperty :String</span></span>

<span data-ttu-id="be02b-203">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="be02b-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="be02b-204">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-204">Type</span></span>

*   <span data-ttu-id="be02b-205">String</span><span class="sxs-lookup"><span data-stu-id="be02b-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="be02b-206">属性：</span><span class="sxs-lookup"><span data-stu-id="be02b-206">Properties:</span></span>

|<span data-ttu-id="be02b-207">名称</span><span class="sxs-lookup"><span data-stu-id="be02b-207">Name</span></span>| <span data-ttu-id="be02b-208">类型</span><span class="sxs-lookup"><span data-stu-id="be02b-208">Type</span></span>| <span data-ttu-id="be02b-209">描述</span><span class="sxs-lookup"><span data-stu-id="be02b-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="be02b-210">字符串</span><span class="sxs-lookup"><span data-stu-id="be02b-210">String</span></span>|<span data-ttu-id="be02b-211">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="be02b-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="be02b-212">String</span><span class="sxs-lookup"><span data-stu-id="be02b-212">String</span></span>|<span data-ttu-id="be02b-213">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="be02b-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be02b-214">Requirements</span><span class="sxs-lookup"><span data-stu-id="be02b-214">Requirements</span></span>

|<span data-ttu-id="be02b-215">要求</span><span class="sxs-lookup"><span data-stu-id="be02b-215">Requirement</span></span>| <span data-ttu-id="be02b-216">值</span><span class="sxs-lookup"><span data-stu-id="be02b-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="be02b-217">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="be02b-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be02b-218">1.0</span><span class="sxs-lookup"><span data-stu-id="be02b-218">1.0</span></span>|
|[<span data-ttu-id="be02b-219">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be02b-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be02b-220">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be02b-220">Compose or Read</span></span>|
