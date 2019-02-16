---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: bbec602680da7914666daf33ed36c45751ae69c6
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068320"
---
# <a name="office"></a><span data-ttu-id="00afa-102">Office</span><span class="sxs-lookup"><span data-stu-id="00afa-102">Office</span></span>

<span data-ttu-id="00afa-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="00afa-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="00afa-105">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-105">Requirements</span></span>

|<span data-ttu-id="00afa-106">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-106">Requirement</span></span>| <span data-ttu-id="00afa-107">值</span><span class="sxs-lookup"><span data-stu-id="00afa-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="00afa-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="00afa-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00afa-109">1.0</span><span class="sxs-lookup"><span data-stu-id="00afa-109">1.0</span></span>|
|[<span data-ttu-id="00afa-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="00afa-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00afa-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="00afa-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="00afa-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="00afa-112">Members and methods</span></span>

| <span data-ttu-id="00afa-113">成员</span><span class="sxs-lookup"><span data-stu-id="00afa-113">Member</span></span> | <span data-ttu-id="00afa-114">类型</span><span class="sxs-lookup"><span data-stu-id="00afa-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="00afa-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="00afa-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="00afa-116">成员</span><span class="sxs-lookup"><span data-stu-id="00afa-116">Member</span></span> |
| [<span data-ttu-id="00afa-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="00afa-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="00afa-118">成员</span><span class="sxs-lookup"><span data-stu-id="00afa-118">Member</span></span> |
| [<span data-ttu-id="00afa-119">EventType</span><span class="sxs-lookup"><span data-stu-id="00afa-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="00afa-120">成员</span><span class="sxs-lookup"><span data-stu-id="00afa-120">Member</span></span> |
| [<span data-ttu-id="00afa-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="00afa-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="00afa-122">成员</span><span class="sxs-lookup"><span data-stu-id="00afa-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="00afa-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="00afa-123">Namespaces</span></span>

<span data-ttu-id="00afa-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="00afa-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="00afa-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="00afa-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="00afa-126">成员</span><span class="sxs-lookup"><span data-stu-id="00afa-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="00afa-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="00afa-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="00afa-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="00afa-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="00afa-129">Type</span><span class="sxs-lookup"><span data-stu-id="00afa-129">Type</span></span>

*   <span data-ttu-id="00afa-130">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="00afa-131">属性：</span><span class="sxs-lookup"><span data-stu-id="00afa-131">Properties:</span></span>

|<span data-ttu-id="00afa-132">名称</span><span class="sxs-lookup"><span data-stu-id="00afa-132">Name</span></span>| <span data-ttu-id="00afa-133">类型</span><span class="sxs-lookup"><span data-stu-id="00afa-133">Type</span></span>| <span data-ttu-id="00afa-134">描述</span><span class="sxs-lookup"><span data-stu-id="00afa-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="00afa-135">String</span><span class="sxs-lookup"><span data-stu-id="00afa-135">String</span></span>|<span data-ttu-id="00afa-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="00afa-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="00afa-137">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-137">String</span></span>|<span data-ttu-id="00afa-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="00afa-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00afa-139">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-139">Requirements</span></span>

|<span data-ttu-id="00afa-140">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-140">Requirement</span></span>| <span data-ttu-id="00afa-141">值</span><span class="sxs-lookup"><span data-stu-id="00afa-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="00afa-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="00afa-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00afa-143">1.0</span><span class="sxs-lookup"><span data-stu-id="00afa-143">1.0</span></span>|
|[<span data-ttu-id="00afa-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="00afa-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00afa-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="00afa-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="00afa-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="00afa-146">CoercionType :String</span></span>

<span data-ttu-id="00afa-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="00afa-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="00afa-148">Type</span><span class="sxs-lookup"><span data-stu-id="00afa-148">Type</span></span>

*   <span data-ttu-id="00afa-149">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="00afa-150">属性：</span><span class="sxs-lookup"><span data-stu-id="00afa-150">Properties:</span></span>

|<span data-ttu-id="00afa-151">名称</span><span class="sxs-lookup"><span data-stu-id="00afa-151">Name</span></span>| <span data-ttu-id="00afa-152">类型</span><span class="sxs-lookup"><span data-stu-id="00afa-152">Type</span></span>| <span data-ttu-id="00afa-153">描述</span><span class="sxs-lookup"><span data-stu-id="00afa-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="00afa-154">String</span><span class="sxs-lookup"><span data-stu-id="00afa-154">String</span></span>|<span data-ttu-id="00afa-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="00afa-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="00afa-156">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-156">String</span></span>|<span data-ttu-id="00afa-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="00afa-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00afa-158">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-158">Requirements</span></span>

|<span data-ttu-id="00afa-159">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-159">Requirement</span></span>| <span data-ttu-id="00afa-160">值</span><span class="sxs-lookup"><span data-stu-id="00afa-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="00afa-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="00afa-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00afa-162">1.0</span><span class="sxs-lookup"><span data-stu-id="00afa-162">1.0</span></span>|
|[<span data-ttu-id="00afa-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="00afa-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00afa-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="00afa-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="00afa-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="00afa-165">EventType :String</span></span>

<span data-ttu-id="00afa-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="00afa-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="00afa-167">Type</span><span class="sxs-lookup"><span data-stu-id="00afa-167">Type</span></span>

*   <span data-ttu-id="00afa-168">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="00afa-169">属性：</span><span class="sxs-lookup"><span data-stu-id="00afa-169">Properties:</span></span>

| <span data-ttu-id="00afa-170">名称</span><span class="sxs-lookup"><span data-stu-id="00afa-170">Name</span></span> | <span data-ttu-id="00afa-171">类型</span><span class="sxs-lookup"><span data-stu-id="00afa-171">Type</span></span> | <span data-ttu-id="00afa-172">描述</span><span class="sxs-lookup"><span data-stu-id="00afa-172">Description</span></span> | <span data-ttu-id="00afa-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="00afa-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="00afa-174">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-174">String</span></span> | <span data-ttu-id="00afa-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="00afa-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="00afa-176">1.7</span><span class="sxs-lookup"><span data-stu-id="00afa-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="00afa-177">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-177">String</span></span> | <span data-ttu-id="00afa-178">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="00afa-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="00afa-179">预览</span><span class="sxs-lookup"><span data-stu-id="00afa-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="00afa-180">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-180">String</span></span> | <span data-ttu-id="00afa-181">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="00afa-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="00afa-182">1.5</span><span class="sxs-lookup"><span data-stu-id="00afa-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="00afa-183">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-183">String</span></span> | <span data-ttu-id="00afa-184">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="00afa-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="00afa-185">预览</span><span class="sxs-lookup"><span data-stu-id="00afa-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="00afa-186">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-186">String</span></span> | <span data-ttu-id="00afa-187">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="00afa-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="00afa-188">1.7</span><span class="sxs-lookup"><span data-stu-id="00afa-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="00afa-189">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-189">String</span></span> | <span data-ttu-id="00afa-190">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="00afa-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="00afa-191">1.7</span><span class="sxs-lookup"><span data-stu-id="00afa-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00afa-192">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-192">Requirements</span></span>

|<span data-ttu-id="00afa-193">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-193">Requirement</span></span>| <span data-ttu-id="00afa-194">值</span><span class="sxs-lookup"><span data-stu-id="00afa-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="00afa-195">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="00afa-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00afa-196">1.5</span><span class="sxs-lookup"><span data-stu-id="00afa-196">1.5</span></span> |
|[<span data-ttu-id="00afa-197">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="00afa-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00afa-198">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="00afa-198">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="00afa-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="00afa-199">SourceProperty :String</span></span>

<span data-ttu-id="00afa-200">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="00afa-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="00afa-201">Type</span><span class="sxs-lookup"><span data-stu-id="00afa-201">Type</span></span>

*   <span data-ttu-id="00afa-202">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="00afa-203">属性：</span><span class="sxs-lookup"><span data-stu-id="00afa-203">Properties:</span></span>

|<span data-ttu-id="00afa-204">名称</span><span class="sxs-lookup"><span data-stu-id="00afa-204">Name</span></span>| <span data-ttu-id="00afa-205">类型</span><span class="sxs-lookup"><span data-stu-id="00afa-205">Type</span></span>| <span data-ttu-id="00afa-206">描述</span><span class="sxs-lookup"><span data-stu-id="00afa-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="00afa-207">字符串</span><span class="sxs-lookup"><span data-stu-id="00afa-207">String</span></span>|<span data-ttu-id="00afa-208">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="00afa-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="00afa-209">String</span><span class="sxs-lookup"><span data-stu-id="00afa-209">String</span></span>|<span data-ttu-id="00afa-210">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="00afa-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00afa-211">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-211">Requirements</span></span>

|<span data-ttu-id="00afa-212">要求</span><span class="sxs-lookup"><span data-stu-id="00afa-212">Requirement</span></span>| <span data-ttu-id="00afa-213">值</span><span class="sxs-lookup"><span data-stu-id="00afa-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="00afa-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="00afa-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00afa-215">1.0</span><span class="sxs-lookup"><span data-stu-id="00afa-215">1.0</span></span>|
|[<span data-ttu-id="00afa-216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="00afa-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="00afa-217">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="00afa-217">Compose or Read</span></span>|
