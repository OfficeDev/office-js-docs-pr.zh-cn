---
title: Office 命名空间-要求集1。7
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 8d22ce8400916dffe12a15bba35f70ceca4db510
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695866"
---
# <a name="office"></a><span data-ttu-id="42f6f-102">Office</span><span class="sxs-lookup"><span data-stu-id="42f6f-102">Office</span></span>

<span data-ttu-id="42f6f-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="42f6f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="42f6f-105">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-105">Requirements</span></span>

|<span data-ttu-id="42f6f-106">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-106">Requirement</span></span>| <span data-ttu-id="42f6f-107">值</span><span class="sxs-lookup"><span data-stu-id="42f6f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="42f6f-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42f6f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42f6f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="42f6f-109">1.0</span></span>|
|[<span data-ttu-id="42f6f-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42f6f-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42f6f-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42f6f-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="42f6f-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="42f6f-112">Members and methods</span></span>

| <span data-ttu-id="42f6f-113">成员</span><span class="sxs-lookup"><span data-stu-id="42f6f-113">Member</span></span> | <span data-ttu-id="42f6f-114">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="42f6f-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="42f6f-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="42f6f-116">Member</span><span class="sxs-lookup"><span data-stu-id="42f6f-116">Member</span></span> |
| [<span data-ttu-id="42f6f-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="42f6f-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="42f6f-118">Member</span><span class="sxs-lookup"><span data-stu-id="42f6f-118">Member</span></span> |
| [<span data-ttu-id="42f6f-119">EventType</span><span class="sxs-lookup"><span data-stu-id="42f6f-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="42f6f-120">Member</span><span class="sxs-lookup"><span data-stu-id="42f6f-120">Member</span></span> |
| [<span data-ttu-id="42f6f-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="42f6f-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="42f6f-122">成员</span><span class="sxs-lookup"><span data-stu-id="42f6f-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="42f6f-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="42f6f-123">Namespaces</span></span>

<span data-ttu-id="42f6f-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="42f6f-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="42f6f-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): 包含多个`ItemType`枚举, 例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="42f6f-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="42f6f-126">Members</span><span class="sxs-lookup"><span data-stu-id="42f6f-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="42f6f-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="42f6f-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="42f6f-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="42f6f-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="42f6f-129">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-129">Type</span></span>

*   <span data-ttu-id="42f6f-130">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="42f6f-131">属性：</span><span class="sxs-lookup"><span data-stu-id="42f6f-131">Properties:</span></span>

|<span data-ttu-id="42f6f-132">名称</span><span class="sxs-lookup"><span data-stu-id="42f6f-132">Name</span></span>| <span data-ttu-id="42f6f-133">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-133">Type</span></span>| <span data-ttu-id="42f6f-134">说明</span><span class="sxs-lookup"><span data-stu-id="42f6f-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="42f6f-135">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-135">String</span></span>|<span data-ttu-id="42f6f-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="42f6f-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="42f6f-137">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-137">String</span></span>|<span data-ttu-id="42f6f-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="42f6f-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42f6f-139">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-139">Requirements</span></span>

|<span data-ttu-id="42f6f-140">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-140">Requirement</span></span>| <span data-ttu-id="42f6f-141">值</span><span class="sxs-lookup"><span data-stu-id="42f6f-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="42f6f-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42f6f-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42f6f-143">1.0</span><span class="sxs-lookup"><span data-stu-id="42f6f-143">1.0</span></span>|
|[<span data-ttu-id="42f6f-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42f6f-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42f6f-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42f6f-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="42f6f-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="42f6f-146">CoercionType: String</span></span>

<span data-ttu-id="42f6f-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="42f6f-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="42f6f-148">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-148">Type</span></span>

*   <span data-ttu-id="42f6f-149">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="42f6f-150">属性：</span><span class="sxs-lookup"><span data-stu-id="42f6f-150">Properties:</span></span>

|<span data-ttu-id="42f6f-151">名称</span><span class="sxs-lookup"><span data-stu-id="42f6f-151">Name</span></span>| <span data-ttu-id="42f6f-152">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-152">Type</span></span>| <span data-ttu-id="42f6f-153">说明</span><span class="sxs-lookup"><span data-stu-id="42f6f-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="42f6f-154">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-154">String</span></span>|<span data-ttu-id="42f6f-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="42f6f-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="42f6f-156">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-156">String</span></span>|<span data-ttu-id="42f6f-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="42f6f-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42f6f-158">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-158">Requirements</span></span>

|<span data-ttu-id="42f6f-159">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-159">Requirement</span></span>| <span data-ttu-id="42f6f-160">值</span><span class="sxs-lookup"><span data-stu-id="42f6f-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="42f6f-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42f6f-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42f6f-162">1.0</span><span class="sxs-lookup"><span data-stu-id="42f6f-162">1.0</span></span>|
|[<span data-ttu-id="42f6f-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42f6f-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42f6f-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42f6f-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="42f6f-165">事件类型: String</span><span class="sxs-lookup"><span data-stu-id="42f6f-165">EventType: String</span></span>

<span data-ttu-id="42f6f-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="42f6f-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="42f6f-167">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-167">Type</span></span>

*   <span data-ttu-id="42f6f-168">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="42f6f-169">属性：</span><span class="sxs-lookup"><span data-stu-id="42f6f-169">Properties:</span></span>

| <span data-ttu-id="42f6f-170">名称</span><span class="sxs-lookup"><span data-stu-id="42f6f-170">Name</span></span> | <span data-ttu-id="42f6f-171">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-171">Type</span></span> | <span data-ttu-id="42f6f-172">说明</span><span class="sxs-lookup"><span data-stu-id="42f6f-172">Description</span></span> | <span data-ttu-id="42f6f-173">最低要求集</span><span class="sxs-lookup"><span data-stu-id="42f6f-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="42f6f-174">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-174">String</span></span> | <span data-ttu-id="42f6f-175">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="42f6f-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="42f6f-176">1.7</span><span class="sxs-lookup"><span data-stu-id="42f6f-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="42f6f-177">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-177">String</span></span> | <span data-ttu-id="42f6f-178">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="42f6f-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="42f6f-179">1.5</span><span class="sxs-lookup"><span data-stu-id="42f6f-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="42f6f-180">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-180">String</span></span> | <span data-ttu-id="42f6f-181">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="42f6f-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="42f6f-182">1.7</span><span class="sxs-lookup"><span data-stu-id="42f6f-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="42f6f-183">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-183">String</span></span> | <span data-ttu-id="42f6f-184">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="42f6f-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="42f6f-185">1.7</span><span class="sxs-lookup"><span data-stu-id="42f6f-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="42f6f-186">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-186">Requirements</span></span>

|<span data-ttu-id="42f6f-187">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-187">Requirement</span></span>| <span data-ttu-id="42f6f-188">值</span><span class="sxs-lookup"><span data-stu-id="42f6f-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="42f6f-189">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42f6f-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42f6f-190">1.5</span><span class="sxs-lookup"><span data-stu-id="42f6f-190">1.5</span></span> |
|[<span data-ttu-id="42f6f-191">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42f6f-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42f6f-192">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42f6f-192">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="42f6f-193">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="42f6f-193">SourceProperty: String</span></span>

<span data-ttu-id="42f6f-194">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="42f6f-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="42f6f-195">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-195">Type</span></span>

*   <span data-ttu-id="42f6f-196">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="42f6f-197">属性：</span><span class="sxs-lookup"><span data-stu-id="42f6f-197">Properties:</span></span>

|<span data-ttu-id="42f6f-198">名称</span><span class="sxs-lookup"><span data-stu-id="42f6f-198">Name</span></span>| <span data-ttu-id="42f6f-199">类型</span><span class="sxs-lookup"><span data-stu-id="42f6f-199">Type</span></span>| <span data-ttu-id="42f6f-200">说明</span><span class="sxs-lookup"><span data-stu-id="42f6f-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="42f6f-201">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-201">String</span></span>|<span data-ttu-id="42f6f-202">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="42f6f-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="42f6f-203">String</span><span class="sxs-lookup"><span data-stu-id="42f6f-203">String</span></span>|<span data-ttu-id="42f6f-204">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="42f6f-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42f6f-205">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-205">Requirements</span></span>

|<span data-ttu-id="42f6f-206">要求</span><span class="sxs-lookup"><span data-stu-id="42f6f-206">Requirement</span></span>| <span data-ttu-id="42f6f-207">值</span><span class="sxs-lookup"><span data-stu-id="42f6f-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="42f6f-208">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42f6f-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42f6f-209">1.0</span><span class="sxs-lookup"><span data-stu-id="42f6f-209">1.0</span></span>|
|[<span data-ttu-id="42f6f-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42f6f-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42f6f-211">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42f6f-211">Compose or Read</span></span>|
