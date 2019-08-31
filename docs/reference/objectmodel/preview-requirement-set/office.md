---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: eb8ff0a755c1908d7b96438f96386056cc16b24f
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696433"
---
# <a name="office"></a><span data-ttu-id="7f37f-102">Office</span><span class="sxs-lookup"><span data-stu-id="7f37f-102">Office</span></span>

<span data-ttu-id="7f37f-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="7f37f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f37f-105">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-105">Requirements</span></span>

|<span data-ttu-id="7f37f-106">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-106">Requirement</span></span>| <span data-ttu-id="7f37f-107">值</span><span class="sxs-lookup"><span data-stu-id="7f37f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f37f-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f37f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f37f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7f37f-109">1.0</span></span>|
|[<span data-ttu-id="7f37f-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f37f-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f37f-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f37f-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7f37f-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="7f37f-112">Members and methods</span></span>

| <span data-ttu-id="7f37f-113">成员</span><span class="sxs-lookup"><span data-stu-id="7f37f-113">Member</span></span> | <span data-ttu-id="7f37f-114">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7f37f-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="7f37f-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="7f37f-116">Member</span><span class="sxs-lookup"><span data-stu-id="7f37f-116">Member</span></span> |
| [<span data-ttu-id="7f37f-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="7f37f-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="7f37f-118">Member</span><span class="sxs-lookup"><span data-stu-id="7f37f-118">Member</span></span> |
| [<span data-ttu-id="7f37f-119">EventType</span><span class="sxs-lookup"><span data-stu-id="7f37f-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="7f37f-120">Member</span><span class="sxs-lookup"><span data-stu-id="7f37f-120">Member</span></span> |
| [<span data-ttu-id="7f37f-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="7f37f-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="7f37f-122">成员</span><span class="sxs-lookup"><span data-stu-id="7f37f-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7f37f-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="7f37f-123">Namespaces</span></span>

<span data-ttu-id="7f37f-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="7f37f-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="7f37f-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): 包含多个`ItemType`枚举, 例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="7f37f-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="7f37f-126">Members</span><span class="sxs-lookup"><span data-stu-id="7f37f-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="7f37f-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="7f37f-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="7f37f-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="7f37f-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7f37f-129">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-129">Type</span></span>

*   <span data-ttu-id="7f37f-130">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7f37f-131">属性：</span><span class="sxs-lookup"><span data-stu-id="7f37f-131">Properties:</span></span>

|<span data-ttu-id="7f37f-132">名称</span><span class="sxs-lookup"><span data-stu-id="7f37f-132">Name</span></span>| <span data-ttu-id="7f37f-133">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-133">Type</span></span>| <span data-ttu-id="7f37f-134">说明</span><span class="sxs-lookup"><span data-stu-id="7f37f-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7f37f-135">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-135">String</span></span>|<span data-ttu-id="7f37f-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="7f37f-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7f37f-137">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-137">String</span></span>|<span data-ttu-id="7f37f-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="7f37f-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f37f-139">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-139">Requirements</span></span>

|<span data-ttu-id="7f37f-140">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-140">Requirement</span></span>| <span data-ttu-id="7f37f-141">值</span><span class="sxs-lookup"><span data-stu-id="7f37f-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f37f-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f37f-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f37f-143">1.0</span><span class="sxs-lookup"><span data-stu-id="7f37f-143">1.0</span></span>|
|[<span data-ttu-id="7f37f-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f37f-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f37f-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f37f-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="7f37f-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="7f37f-146">CoercionType: String</span></span>

<span data-ttu-id="7f37f-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="7f37f-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7f37f-148">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-148">Type</span></span>

*   <span data-ttu-id="7f37f-149">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7f37f-150">属性：</span><span class="sxs-lookup"><span data-stu-id="7f37f-150">Properties:</span></span>

|<span data-ttu-id="7f37f-151">名称</span><span class="sxs-lookup"><span data-stu-id="7f37f-151">Name</span></span>| <span data-ttu-id="7f37f-152">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-152">Type</span></span>| <span data-ttu-id="7f37f-153">说明</span><span class="sxs-lookup"><span data-stu-id="7f37f-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7f37f-154">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-154">String</span></span>|<span data-ttu-id="7f37f-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="7f37f-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7f37f-156">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-156">String</span></span>|<span data-ttu-id="7f37f-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="7f37f-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f37f-158">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-158">Requirements</span></span>

|<span data-ttu-id="7f37f-159">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-159">Requirement</span></span>| <span data-ttu-id="7f37f-160">值</span><span class="sxs-lookup"><span data-stu-id="7f37f-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f37f-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f37f-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f37f-162">1.0</span><span class="sxs-lookup"><span data-stu-id="7f37f-162">1.0</span></span>|
|[<span data-ttu-id="7f37f-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f37f-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f37f-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f37f-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="7f37f-165">事件类型: String</span><span class="sxs-lookup"><span data-stu-id="7f37f-165">EventType: String</span></span>

<span data-ttu-id="7f37f-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="7f37f-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="7f37f-167">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-167">Type</span></span>

*   <span data-ttu-id="7f37f-168">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7f37f-169">属性：</span><span class="sxs-lookup"><span data-stu-id="7f37f-169">Properties:</span></span>

| <span data-ttu-id="7f37f-170">名称</span><span class="sxs-lookup"><span data-stu-id="7f37f-170">Name</span></span> | <span data-ttu-id="7f37f-171">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-171">Type</span></span> | <span data-ttu-id="7f37f-172">说明</span><span class="sxs-lookup"><span data-stu-id="7f37f-172">Description</span></span> | <span data-ttu-id="7f37f-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="7f37f-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="7f37f-174">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-174">String</span></span> | <span data-ttu-id="7f37f-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="7f37f-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="7f37f-176">1.7</span><span class="sxs-lookup"><span data-stu-id="7f37f-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="7f37f-177">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-177">String</span></span> | <span data-ttu-id="7f37f-178">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="7f37f-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="7f37f-179">预览</span><span class="sxs-lookup"><span data-stu-id="7f37f-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="7f37f-180">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-180">String</span></span> | <span data-ttu-id="7f37f-181">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="7f37f-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="7f37f-182">预览</span><span class="sxs-lookup"><span data-stu-id="7f37f-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="7f37f-183">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-183">String</span></span> | <span data-ttu-id="7f37f-184">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="7f37f-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="7f37f-185">1.5</span><span class="sxs-lookup"><span data-stu-id="7f37f-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="7f37f-186">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-186">String</span></span> | <span data-ttu-id="7f37f-187">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="7f37f-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="7f37f-188">预览</span><span class="sxs-lookup"><span data-stu-id="7f37f-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="7f37f-189">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-189">String</span></span> | <span data-ttu-id="7f37f-190">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="7f37f-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="7f37f-191">1.7</span><span class="sxs-lookup"><span data-stu-id="7f37f-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="7f37f-192">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-192">String</span></span> | <span data-ttu-id="7f37f-193">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="7f37f-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="7f37f-194">1.7</span><span class="sxs-lookup"><span data-stu-id="7f37f-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7f37f-195">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-195">Requirements</span></span>

|<span data-ttu-id="7f37f-196">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-196">Requirement</span></span>| <span data-ttu-id="7f37f-197">值</span><span class="sxs-lookup"><span data-stu-id="7f37f-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f37f-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f37f-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f37f-199">1.5</span><span class="sxs-lookup"><span data-stu-id="7f37f-199">1.5</span></span> |
|[<span data-ttu-id="7f37f-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f37f-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f37f-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f37f-201">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="7f37f-202">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="7f37f-202">SourceProperty: String</span></span>

<span data-ttu-id="7f37f-203">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="7f37f-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7f37f-204">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-204">Type</span></span>

*   <span data-ttu-id="7f37f-205">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7f37f-206">属性：</span><span class="sxs-lookup"><span data-stu-id="7f37f-206">Properties:</span></span>

|<span data-ttu-id="7f37f-207">名称</span><span class="sxs-lookup"><span data-stu-id="7f37f-207">Name</span></span>| <span data-ttu-id="7f37f-208">类型</span><span class="sxs-lookup"><span data-stu-id="7f37f-208">Type</span></span>| <span data-ttu-id="7f37f-209">说明</span><span class="sxs-lookup"><span data-stu-id="7f37f-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7f37f-210">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-210">String</span></span>|<span data-ttu-id="7f37f-211">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="7f37f-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7f37f-212">String</span><span class="sxs-lookup"><span data-stu-id="7f37f-212">String</span></span>|<span data-ttu-id="7f37f-213">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="7f37f-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f37f-214">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-214">Requirements</span></span>

|<span data-ttu-id="7f37f-215">要求</span><span class="sxs-lookup"><span data-stu-id="7f37f-215">Requirement</span></span>| <span data-ttu-id="7f37f-216">值</span><span class="sxs-lookup"><span data-stu-id="7f37f-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f37f-217">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7f37f-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f37f-218">1.0</span><span class="sxs-lookup"><span data-stu-id="7f37f-218">1.0</span></span>|
|[<span data-ttu-id="7f37f-219">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7f37f-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f37f-220">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7f37f-220">Compose or Read</span></span>|
