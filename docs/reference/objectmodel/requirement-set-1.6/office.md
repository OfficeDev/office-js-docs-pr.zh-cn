---
title: Office 命名空间-要求集1。6
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: ae764e8cda2b3f14e33b883d054379db7b37a687
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695999"
---
# <a name="office"></a><span data-ttu-id="bf340-102">Office</span><span class="sxs-lookup"><span data-stu-id="bf340-102">Office</span></span>

<span data-ttu-id="bf340-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="bf340-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf340-105">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-105">Requirements</span></span>

|<span data-ttu-id="bf340-106">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-106">Requirement</span></span>| <span data-ttu-id="bf340-107">值</span><span class="sxs-lookup"><span data-stu-id="bf340-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf340-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bf340-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf340-109">1.0</span><span class="sxs-lookup"><span data-stu-id="bf340-109">1.0</span></span>|
|[<span data-ttu-id="bf340-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bf340-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf340-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bf340-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bf340-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="bf340-112">Members and methods</span></span>

| <span data-ttu-id="bf340-113">成员</span><span class="sxs-lookup"><span data-stu-id="bf340-113">Member</span></span> | <span data-ttu-id="bf340-114">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bf340-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="bf340-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="bf340-116">Member</span><span class="sxs-lookup"><span data-stu-id="bf340-116">Member</span></span> |
| [<span data-ttu-id="bf340-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="bf340-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="bf340-118">Member</span><span class="sxs-lookup"><span data-stu-id="bf340-118">Member</span></span> |
| [<span data-ttu-id="bf340-119">EventType</span><span class="sxs-lookup"><span data-stu-id="bf340-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="bf340-120">Member</span><span class="sxs-lookup"><span data-stu-id="bf340-120">Member</span></span> |
| [<span data-ttu-id="bf340-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="bf340-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="bf340-122">成员</span><span class="sxs-lookup"><span data-stu-id="bf340-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="bf340-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="bf340-123">Namespaces</span></span>

<span data-ttu-id="bf340-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="bf340-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="bf340-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): 包含多个`ItemType`枚举, 例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="bf340-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="bf340-126">Members</span><span class="sxs-lookup"><span data-stu-id="bf340-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="bf340-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="bf340-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="bf340-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="bf340-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="bf340-129">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-129">Type</span></span>

*   <span data-ttu-id="bf340-130">String</span><span class="sxs-lookup"><span data-stu-id="bf340-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bf340-131">属性：</span><span class="sxs-lookup"><span data-stu-id="bf340-131">Properties:</span></span>

|<span data-ttu-id="bf340-132">名称</span><span class="sxs-lookup"><span data-stu-id="bf340-132">Name</span></span>| <span data-ttu-id="bf340-133">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-133">Type</span></span>| <span data-ttu-id="bf340-134">说明</span><span class="sxs-lookup"><span data-stu-id="bf340-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="bf340-135">String</span><span class="sxs-lookup"><span data-stu-id="bf340-135">String</span></span>|<span data-ttu-id="bf340-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="bf340-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="bf340-137">String</span><span class="sxs-lookup"><span data-stu-id="bf340-137">String</span></span>|<span data-ttu-id="bf340-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="bf340-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf340-139">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-139">Requirements</span></span>

|<span data-ttu-id="bf340-140">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-140">Requirement</span></span>| <span data-ttu-id="bf340-141">值</span><span class="sxs-lookup"><span data-stu-id="bf340-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf340-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bf340-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf340-143">1.0</span><span class="sxs-lookup"><span data-stu-id="bf340-143">1.0</span></span>|
|[<span data-ttu-id="bf340-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bf340-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf340-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bf340-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="bf340-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="bf340-146">CoercionType: String</span></span>

<span data-ttu-id="bf340-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="bf340-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bf340-148">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-148">Type</span></span>

*   <span data-ttu-id="bf340-149">String</span><span class="sxs-lookup"><span data-stu-id="bf340-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bf340-150">属性：</span><span class="sxs-lookup"><span data-stu-id="bf340-150">Properties:</span></span>

|<span data-ttu-id="bf340-151">名称</span><span class="sxs-lookup"><span data-stu-id="bf340-151">Name</span></span>| <span data-ttu-id="bf340-152">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-152">Type</span></span>| <span data-ttu-id="bf340-153">说明</span><span class="sxs-lookup"><span data-stu-id="bf340-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="bf340-154">String</span><span class="sxs-lookup"><span data-stu-id="bf340-154">String</span></span>|<span data-ttu-id="bf340-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="bf340-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="bf340-156">String</span><span class="sxs-lookup"><span data-stu-id="bf340-156">String</span></span>|<span data-ttu-id="bf340-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="bf340-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf340-158">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-158">Requirements</span></span>

|<span data-ttu-id="bf340-159">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-159">Requirement</span></span>| <span data-ttu-id="bf340-160">值</span><span class="sxs-lookup"><span data-stu-id="bf340-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf340-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bf340-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf340-162">1.0</span><span class="sxs-lookup"><span data-stu-id="bf340-162">1.0</span></span>|
|[<span data-ttu-id="bf340-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bf340-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf340-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bf340-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="bf340-165">事件类型: String</span><span class="sxs-lookup"><span data-stu-id="bf340-165">EventType: String</span></span>

<span data-ttu-id="bf340-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="bf340-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="bf340-167">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-167">Type</span></span>

*   <span data-ttu-id="bf340-168">String</span><span class="sxs-lookup"><span data-stu-id="bf340-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bf340-169">属性：</span><span class="sxs-lookup"><span data-stu-id="bf340-169">Properties:</span></span>

| <span data-ttu-id="bf340-170">名称</span><span class="sxs-lookup"><span data-stu-id="bf340-170">Name</span></span> | <span data-ttu-id="bf340-171">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-171">Type</span></span> | <span data-ttu-id="bf340-172">说明</span><span class="sxs-lookup"><span data-stu-id="bf340-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="bf340-173">String</span><span class="sxs-lookup"><span data-stu-id="bf340-173">String</span></span> | <span data-ttu-id="bf340-174">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="bf340-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bf340-175">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-175">Requirements</span></span>

|<span data-ttu-id="bf340-176">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-176">Requirement</span></span>| <span data-ttu-id="bf340-177">值</span><span class="sxs-lookup"><span data-stu-id="bf340-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf340-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bf340-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf340-179">1.5</span><span class="sxs-lookup"><span data-stu-id="bf340-179">1.5</span></span> |
|[<span data-ttu-id="bf340-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bf340-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf340-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bf340-181">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="bf340-182">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="bf340-182">SourceProperty: String</span></span>

<span data-ttu-id="bf340-183">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="bf340-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bf340-184">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-184">Type</span></span>

*   <span data-ttu-id="bf340-185">String</span><span class="sxs-lookup"><span data-stu-id="bf340-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bf340-186">属性：</span><span class="sxs-lookup"><span data-stu-id="bf340-186">Properties:</span></span>

|<span data-ttu-id="bf340-187">名称</span><span class="sxs-lookup"><span data-stu-id="bf340-187">Name</span></span>| <span data-ttu-id="bf340-188">类型</span><span class="sxs-lookup"><span data-stu-id="bf340-188">Type</span></span>| <span data-ttu-id="bf340-189">说明</span><span class="sxs-lookup"><span data-stu-id="bf340-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="bf340-190">String</span><span class="sxs-lookup"><span data-stu-id="bf340-190">String</span></span>|<span data-ttu-id="bf340-191">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="bf340-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="bf340-192">String</span><span class="sxs-lookup"><span data-stu-id="bf340-192">String</span></span>|<span data-ttu-id="bf340-193">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="bf340-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf340-194">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-194">Requirements</span></span>

|<span data-ttu-id="bf340-195">要求</span><span class="sxs-lookup"><span data-stu-id="bf340-195">Requirement</span></span>| <span data-ttu-id="bf340-196">值</span><span class="sxs-lookup"><span data-stu-id="bf340-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf340-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bf340-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bf340-198">1.0</span><span class="sxs-lookup"><span data-stu-id="bf340-198">1.0</span></span>|
|[<span data-ttu-id="bf340-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bf340-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf340-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bf340-200">Compose or Read</span></span>|
