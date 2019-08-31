---
title: Office 命名空间-要求集1。5
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 2236dae5421090a571c8cc658cb6f67f2a08d54a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696076"
---
# <a name="office"></a><span data-ttu-id="dc73d-102">Office</span><span class="sxs-lookup"><span data-stu-id="dc73d-102">Office</span></span>

<span data-ttu-id="dc73d-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="dc73d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc73d-105">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-105">Requirements</span></span>

|<span data-ttu-id="dc73d-106">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-106">Requirement</span></span>| <span data-ttu-id="dc73d-107">值</span><span class="sxs-lookup"><span data-stu-id="dc73d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc73d-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc73d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc73d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="dc73d-109">1.0</span></span>|
|[<span data-ttu-id="dc73d-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc73d-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc73d-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc73d-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="dc73d-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="dc73d-112">Members and methods</span></span>

| <span data-ttu-id="dc73d-113">成员</span><span class="sxs-lookup"><span data-stu-id="dc73d-113">Member</span></span> | <span data-ttu-id="dc73d-114">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="dc73d-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="dc73d-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="dc73d-116">Member</span><span class="sxs-lookup"><span data-stu-id="dc73d-116">Member</span></span> |
| [<span data-ttu-id="dc73d-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="dc73d-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="dc73d-118">Member</span><span class="sxs-lookup"><span data-stu-id="dc73d-118">Member</span></span> |
| [<span data-ttu-id="dc73d-119">EventType</span><span class="sxs-lookup"><span data-stu-id="dc73d-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="dc73d-120">Member</span><span class="sxs-lookup"><span data-stu-id="dc73d-120">Member</span></span> |
| [<span data-ttu-id="dc73d-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="dc73d-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="dc73d-122">成员</span><span class="sxs-lookup"><span data-stu-id="dc73d-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="dc73d-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="dc73d-123">Namespaces</span></span>

<span data-ttu-id="dc73d-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="dc73d-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="dc73d-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): 包含多个`ItemType`枚举, 例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="dc73d-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="dc73d-126">Members</span><span class="sxs-lookup"><span data-stu-id="dc73d-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="dc73d-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="dc73d-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="dc73d-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="dc73d-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dc73d-129">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-129">Type</span></span>

*   <span data-ttu-id="dc73d-130">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc73d-131">属性：</span><span class="sxs-lookup"><span data-stu-id="dc73d-131">Properties:</span></span>

|<span data-ttu-id="dc73d-132">名称</span><span class="sxs-lookup"><span data-stu-id="dc73d-132">Name</span></span>| <span data-ttu-id="dc73d-133">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-133">Type</span></span>| <span data-ttu-id="dc73d-134">说明</span><span class="sxs-lookup"><span data-stu-id="dc73d-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dc73d-135">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-135">String</span></span>|<span data-ttu-id="dc73d-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="dc73d-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dc73d-137">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-137">String</span></span>|<span data-ttu-id="dc73d-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="dc73d-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc73d-139">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-139">Requirements</span></span>

|<span data-ttu-id="dc73d-140">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-140">Requirement</span></span>| <span data-ttu-id="dc73d-141">值</span><span class="sxs-lookup"><span data-stu-id="dc73d-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc73d-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc73d-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc73d-143">1.0</span><span class="sxs-lookup"><span data-stu-id="dc73d-143">1.0</span></span>|
|[<span data-ttu-id="dc73d-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc73d-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc73d-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc73d-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="dc73d-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="dc73d-146">CoercionType: String</span></span>

<span data-ttu-id="dc73d-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="dc73d-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc73d-148">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-148">Type</span></span>

*   <span data-ttu-id="dc73d-149">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc73d-150">属性：</span><span class="sxs-lookup"><span data-stu-id="dc73d-150">Properties:</span></span>

|<span data-ttu-id="dc73d-151">名称</span><span class="sxs-lookup"><span data-stu-id="dc73d-151">Name</span></span>| <span data-ttu-id="dc73d-152">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-152">Type</span></span>| <span data-ttu-id="dc73d-153">说明</span><span class="sxs-lookup"><span data-stu-id="dc73d-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dc73d-154">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-154">String</span></span>|<span data-ttu-id="dc73d-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="dc73d-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dc73d-156">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-156">String</span></span>|<span data-ttu-id="dc73d-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="dc73d-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc73d-158">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-158">Requirements</span></span>

|<span data-ttu-id="dc73d-159">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-159">Requirement</span></span>| <span data-ttu-id="dc73d-160">值</span><span class="sxs-lookup"><span data-stu-id="dc73d-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc73d-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc73d-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc73d-162">1.0</span><span class="sxs-lookup"><span data-stu-id="dc73d-162">1.0</span></span>|
|[<span data-ttu-id="dc73d-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc73d-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc73d-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc73d-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="dc73d-165">事件类型: String</span><span class="sxs-lookup"><span data-stu-id="dc73d-165">EventType: String</span></span>

<span data-ttu-id="dc73d-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="dc73d-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="dc73d-167">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-167">Type</span></span>

*   <span data-ttu-id="dc73d-168">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc73d-169">属性：</span><span class="sxs-lookup"><span data-stu-id="dc73d-169">Properties:</span></span>

| <span data-ttu-id="dc73d-170">名称</span><span class="sxs-lookup"><span data-stu-id="dc73d-170">Name</span></span> | <span data-ttu-id="dc73d-171">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-171">Type</span></span> | <span data-ttu-id="dc73d-172">说明</span><span class="sxs-lookup"><span data-stu-id="dc73d-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="dc73d-173">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-173">String</span></span> | <span data-ttu-id="dc73d-174">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="dc73d-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dc73d-175">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-175">Requirements</span></span>

|<span data-ttu-id="dc73d-176">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-176">Requirement</span></span>| <span data-ttu-id="dc73d-177">值</span><span class="sxs-lookup"><span data-stu-id="dc73d-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc73d-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc73d-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc73d-179">1.5</span><span class="sxs-lookup"><span data-stu-id="dc73d-179">1.5</span></span> |
|[<span data-ttu-id="dc73d-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc73d-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc73d-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc73d-181">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="dc73d-182">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="dc73d-182">SourceProperty: String</span></span>

<span data-ttu-id="dc73d-183">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="dc73d-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc73d-184">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-184">Type</span></span>

*   <span data-ttu-id="dc73d-185">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc73d-186">属性：</span><span class="sxs-lookup"><span data-stu-id="dc73d-186">Properties:</span></span>

|<span data-ttu-id="dc73d-187">名称</span><span class="sxs-lookup"><span data-stu-id="dc73d-187">Name</span></span>| <span data-ttu-id="dc73d-188">类型</span><span class="sxs-lookup"><span data-stu-id="dc73d-188">Type</span></span>| <span data-ttu-id="dc73d-189">说明</span><span class="sxs-lookup"><span data-stu-id="dc73d-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dc73d-190">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-190">String</span></span>|<span data-ttu-id="dc73d-191">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="dc73d-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dc73d-192">String</span><span class="sxs-lookup"><span data-stu-id="dc73d-192">String</span></span>|<span data-ttu-id="dc73d-193">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="dc73d-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc73d-194">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-194">Requirements</span></span>

|<span data-ttu-id="dc73d-195">要求</span><span class="sxs-lookup"><span data-stu-id="dc73d-195">Requirement</span></span>| <span data-ttu-id="dc73d-196">值</span><span class="sxs-lookup"><span data-stu-id="dc73d-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc73d-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="dc73d-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc73d-198">1.0</span><span class="sxs-lookup"><span data-stu-id="dc73d-198">1.0</span></span>|
|[<span data-ttu-id="dc73d-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="dc73d-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dc73d-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="dc73d-200">Compose or Read</span></span>|
