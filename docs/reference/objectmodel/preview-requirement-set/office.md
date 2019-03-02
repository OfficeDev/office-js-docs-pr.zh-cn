---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 02/26/2019
localization_priority: Normal
ms.openlocfilehash: 7b27963a85f1dcdaa6f269fce242c45bf1bdd146
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359231"
---
# <a name="office"></a><span data-ttu-id="36d35-102">Office</span><span class="sxs-lookup"><span data-stu-id="36d35-102">Office</span></span>

<span data-ttu-id="36d35-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="36d35-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="36d35-105">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-105">Requirements</span></span>

|<span data-ttu-id="36d35-106">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-106">Requirement</span></span>| <span data-ttu-id="36d35-107">值</span><span class="sxs-lookup"><span data-stu-id="36d35-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="36d35-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="36d35-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36d35-109">1.0</span><span class="sxs-lookup"><span data-stu-id="36d35-109">1.0</span></span>|
|[<span data-ttu-id="36d35-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36d35-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36d35-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36d35-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="36d35-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="36d35-112">Members and methods</span></span>

| <span data-ttu-id="36d35-113">成员</span><span class="sxs-lookup"><span data-stu-id="36d35-113">Member</span></span> | <span data-ttu-id="36d35-114">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="36d35-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="36d35-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="36d35-116">成员</span><span class="sxs-lookup"><span data-stu-id="36d35-116">Member</span></span> |
| [<span data-ttu-id="36d35-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="36d35-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="36d35-118">成员</span><span class="sxs-lookup"><span data-stu-id="36d35-118">Member</span></span> |
| [<span data-ttu-id="36d35-119">EventType</span><span class="sxs-lookup"><span data-stu-id="36d35-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="36d35-120">成员</span><span class="sxs-lookup"><span data-stu-id="36d35-120">Member</span></span> |
| [<span data-ttu-id="36d35-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="36d35-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="36d35-122">成员</span><span class="sxs-lookup"><span data-stu-id="36d35-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="36d35-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="36d35-123">Namespaces</span></span>

<span data-ttu-id="36d35-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="36d35-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="36d35-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="36d35-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="36d35-126">成员</span><span class="sxs-lookup"><span data-stu-id="36d35-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="36d35-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="36d35-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="36d35-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="36d35-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="36d35-129">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-129">Type</span></span>

*   <span data-ttu-id="36d35-130">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="36d35-131">属性：</span><span class="sxs-lookup"><span data-stu-id="36d35-131">Properties:</span></span>

|<span data-ttu-id="36d35-132">名称</span><span class="sxs-lookup"><span data-stu-id="36d35-132">Name</span></span>| <span data-ttu-id="36d35-133">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-133">Type</span></span>| <span data-ttu-id="36d35-134">描述</span><span class="sxs-lookup"><span data-stu-id="36d35-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="36d35-135">String</span><span class="sxs-lookup"><span data-stu-id="36d35-135">String</span></span>|<span data-ttu-id="36d35-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="36d35-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="36d35-137">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-137">String</span></span>|<span data-ttu-id="36d35-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="36d35-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="36d35-139">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-139">Requirements</span></span>

|<span data-ttu-id="36d35-140">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-140">Requirement</span></span>| <span data-ttu-id="36d35-141">值</span><span class="sxs-lookup"><span data-stu-id="36d35-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="36d35-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="36d35-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36d35-143">1.0</span><span class="sxs-lookup"><span data-stu-id="36d35-143">1.0</span></span>|
|[<span data-ttu-id="36d35-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36d35-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36d35-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36d35-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="36d35-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="36d35-146">CoercionType :String</span></span>

<span data-ttu-id="36d35-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="36d35-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="36d35-148">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-148">Type</span></span>

*   <span data-ttu-id="36d35-149">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="36d35-150">属性：</span><span class="sxs-lookup"><span data-stu-id="36d35-150">Properties:</span></span>

|<span data-ttu-id="36d35-151">名称</span><span class="sxs-lookup"><span data-stu-id="36d35-151">Name</span></span>| <span data-ttu-id="36d35-152">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-152">Type</span></span>| <span data-ttu-id="36d35-153">描述</span><span class="sxs-lookup"><span data-stu-id="36d35-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="36d35-154">String</span><span class="sxs-lookup"><span data-stu-id="36d35-154">String</span></span>|<span data-ttu-id="36d35-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="36d35-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="36d35-156">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-156">String</span></span>|<span data-ttu-id="36d35-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="36d35-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="36d35-158">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-158">Requirements</span></span>

|<span data-ttu-id="36d35-159">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-159">Requirement</span></span>| <span data-ttu-id="36d35-160">值</span><span class="sxs-lookup"><span data-stu-id="36d35-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="36d35-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="36d35-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36d35-162">1.0</span><span class="sxs-lookup"><span data-stu-id="36d35-162">1.0</span></span>|
|[<span data-ttu-id="36d35-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36d35-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36d35-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36d35-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="36d35-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="36d35-165">EventType :String</span></span>

<span data-ttu-id="36d35-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="36d35-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="36d35-167">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-167">Type</span></span>

*   <span data-ttu-id="36d35-168">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="36d35-169">属性：</span><span class="sxs-lookup"><span data-stu-id="36d35-169">Properties:</span></span>

| <span data-ttu-id="36d35-170">名称</span><span class="sxs-lookup"><span data-stu-id="36d35-170">Name</span></span> | <span data-ttu-id="36d35-171">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-171">Type</span></span> | <span data-ttu-id="36d35-172">描述</span><span class="sxs-lookup"><span data-stu-id="36d35-172">Description</span></span> | <span data-ttu-id="36d35-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="36d35-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="36d35-174">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-174">String</span></span> | <span data-ttu-id="36d35-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="36d35-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="36d35-176">1.7</span><span class="sxs-lookup"><span data-stu-id="36d35-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="36d35-177">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-177">String</span></span> | <span data-ttu-id="36d35-178">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="36d35-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="36d35-179">预览</span><span class="sxs-lookup"><span data-stu-id="36d35-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="36d35-180">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-180">String</span></span> | <span data-ttu-id="36d35-181">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="36d35-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="36d35-182">预览</span><span class="sxs-lookup"><span data-stu-id="36d35-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="36d35-183">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-183">String</span></span> | <span data-ttu-id="36d35-184">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="36d35-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="36d35-185">1.5</span><span class="sxs-lookup"><span data-stu-id="36d35-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="36d35-186">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-186">String</span></span> | <span data-ttu-id="36d35-187">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="36d35-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="36d35-188">预览</span><span class="sxs-lookup"><span data-stu-id="36d35-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="36d35-189">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-189">String</span></span> | <span data-ttu-id="36d35-190">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="36d35-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="36d35-191">1.7</span><span class="sxs-lookup"><span data-stu-id="36d35-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="36d35-192">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-192">String</span></span> | <span data-ttu-id="36d35-193">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="36d35-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="36d35-194">1.7</span><span class="sxs-lookup"><span data-stu-id="36d35-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="36d35-195">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-195">Requirements</span></span>

|<span data-ttu-id="36d35-196">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-196">Requirement</span></span>| <span data-ttu-id="36d35-197">值</span><span class="sxs-lookup"><span data-stu-id="36d35-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="36d35-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="36d35-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36d35-199">1.5</span><span class="sxs-lookup"><span data-stu-id="36d35-199">1.5</span></span> |
|[<span data-ttu-id="36d35-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36d35-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36d35-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36d35-201">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="36d35-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="36d35-202">SourceProperty :String</span></span>

<span data-ttu-id="36d35-203">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="36d35-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="36d35-204">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-204">Type</span></span>

*   <span data-ttu-id="36d35-205">String</span><span class="sxs-lookup"><span data-stu-id="36d35-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="36d35-206">属性：</span><span class="sxs-lookup"><span data-stu-id="36d35-206">Properties:</span></span>

|<span data-ttu-id="36d35-207">名称</span><span class="sxs-lookup"><span data-stu-id="36d35-207">Name</span></span>| <span data-ttu-id="36d35-208">类型</span><span class="sxs-lookup"><span data-stu-id="36d35-208">Type</span></span>| <span data-ttu-id="36d35-209">说明</span><span class="sxs-lookup"><span data-stu-id="36d35-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="36d35-210">String</span><span class="sxs-lookup"><span data-stu-id="36d35-210">String</span></span>|<span data-ttu-id="36d35-211">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="36d35-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="36d35-212">字符串</span><span class="sxs-lookup"><span data-stu-id="36d35-212">String</span></span>|<span data-ttu-id="36d35-213">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="36d35-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="36d35-214">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-214">Requirements</span></span>

|<span data-ttu-id="36d35-215">要求</span><span class="sxs-lookup"><span data-stu-id="36d35-215">Requirement</span></span>| <span data-ttu-id="36d35-216">值</span><span class="sxs-lookup"><span data-stu-id="36d35-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="36d35-217">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="36d35-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="36d35-218">1.0</span><span class="sxs-lookup"><span data-stu-id="36d35-218">1.0</span></span>|
|[<span data-ttu-id="36d35-219">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="36d35-219">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="36d35-220">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="36d35-220">Compose or Read</span></span>|
