---
title: Office 命名空间-要求集1。6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dde96f48863459da5072d6b4864169f198264133
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450371"
---
# <a name="office"></a><span data-ttu-id="f13a6-102">Office</span><span class="sxs-lookup"><span data-stu-id="f13a6-102">Office</span></span>

<span data-ttu-id="f13a6-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="f13a6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f13a6-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="f13a6-105">Requirements</span></span>

|<span data-ttu-id="f13a6-106">要求</span><span class="sxs-lookup"><span data-stu-id="f13a6-106">Requirement</span></span>| <span data-ttu-id="f13a6-107">值</span><span class="sxs-lookup"><span data-stu-id="f13a6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f13a6-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f13a6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f13a6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f13a6-109">1.0</span></span>|
|[<span data-ttu-id="f13a6-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f13a6-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f13a6-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f13a6-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f13a6-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="f13a6-112">Members and methods</span></span>

| <span data-ttu-id="f13a6-113">成员</span><span class="sxs-lookup"><span data-stu-id="f13a6-113">Member</span></span> | <span data-ttu-id="f13a6-114">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f13a6-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f13a6-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f13a6-116">Member</span><span class="sxs-lookup"><span data-stu-id="f13a6-116">Member</span></span> |
| [<span data-ttu-id="f13a6-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f13a6-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f13a6-118">Member</span><span class="sxs-lookup"><span data-stu-id="f13a6-118">Member</span></span> |
| [<span data-ttu-id="f13a6-119">EventType</span><span class="sxs-lookup"><span data-stu-id="f13a6-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f13a6-120">Member</span><span class="sxs-lookup"><span data-stu-id="f13a6-120">Member</span></span> |
| [<span data-ttu-id="f13a6-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f13a6-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f13a6-122">成员</span><span class="sxs-lookup"><span data-stu-id="f13a6-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f13a6-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="f13a6-123">Namespaces</span></span>

<span data-ttu-id="f13a6-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="f13a6-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f13a6-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="f13a6-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f13a6-126">成员</span><span class="sxs-lookup"><span data-stu-id="f13a6-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f13a6-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f13a6-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="f13a6-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="f13a6-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f13a6-129">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-129">Type</span></span>

*   <span data-ttu-id="f13a6-130">String</span><span class="sxs-lookup"><span data-stu-id="f13a6-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f13a6-131">属性：</span><span class="sxs-lookup"><span data-stu-id="f13a6-131">Properties:</span></span>

|<span data-ttu-id="f13a6-132">名称</span><span class="sxs-lookup"><span data-stu-id="f13a6-132">Name</span></span>| <span data-ttu-id="f13a6-133">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-133">Type</span></span>| <span data-ttu-id="f13a6-134">说明</span><span class="sxs-lookup"><span data-stu-id="f13a6-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f13a6-135">字符串</span><span class="sxs-lookup"><span data-stu-id="f13a6-135">String</span></span>|<span data-ttu-id="f13a6-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="f13a6-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f13a6-137">字符串</span><span class="sxs-lookup"><span data-stu-id="f13a6-137">String</span></span>|<span data-ttu-id="f13a6-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="f13a6-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f13a6-139">Requirements</span><span class="sxs-lookup"><span data-stu-id="f13a6-139">Requirements</span></span>

|<span data-ttu-id="f13a6-140">要求</span><span class="sxs-lookup"><span data-stu-id="f13a6-140">Requirement</span></span>| <span data-ttu-id="f13a6-141">值</span><span class="sxs-lookup"><span data-stu-id="f13a6-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="f13a6-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f13a6-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f13a6-143">1.0</span><span class="sxs-lookup"><span data-stu-id="f13a6-143">1.0</span></span>|
|[<span data-ttu-id="f13a6-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f13a6-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f13a6-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f13a6-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="f13a6-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f13a6-146">CoercionType :String</span></span>

<span data-ttu-id="f13a6-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="f13a6-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f13a6-148">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-148">Type</span></span>

*   <span data-ttu-id="f13a6-149">String</span><span class="sxs-lookup"><span data-stu-id="f13a6-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f13a6-150">属性：</span><span class="sxs-lookup"><span data-stu-id="f13a6-150">Properties:</span></span>

|<span data-ttu-id="f13a6-151">名称</span><span class="sxs-lookup"><span data-stu-id="f13a6-151">Name</span></span>| <span data-ttu-id="f13a6-152">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-152">Type</span></span>| <span data-ttu-id="f13a6-153">说明</span><span class="sxs-lookup"><span data-stu-id="f13a6-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f13a6-154">字符串</span><span class="sxs-lookup"><span data-stu-id="f13a6-154">String</span></span>|<span data-ttu-id="f13a6-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="f13a6-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f13a6-156">字符串</span><span class="sxs-lookup"><span data-stu-id="f13a6-156">String</span></span>|<span data-ttu-id="f13a6-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="f13a6-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f13a6-158">Requirements</span><span class="sxs-lookup"><span data-stu-id="f13a6-158">Requirements</span></span>

|<span data-ttu-id="f13a6-159">要求</span><span class="sxs-lookup"><span data-stu-id="f13a6-159">Requirement</span></span>| <span data-ttu-id="f13a6-160">值</span><span class="sxs-lookup"><span data-stu-id="f13a6-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="f13a6-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f13a6-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f13a6-162">1.0</span><span class="sxs-lookup"><span data-stu-id="f13a6-162">1.0</span></span>|
|[<span data-ttu-id="f13a6-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f13a6-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f13a6-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f13a6-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="f13a6-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="f13a6-165">EventType :String</span></span>

<span data-ttu-id="f13a6-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="f13a6-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f13a6-167">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-167">Type</span></span>

*   <span data-ttu-id="f13a6-168">String</span><span class="sxs-lookup"><span data-stu-id="f13a6-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f13a6-169">属性：</span><span class="sxs-lookup"><span data-stu-id="f13a6-169">Properties:</span></span>

| <span data-ttu-id="f13a6-170">名称</span><span class="sxs-lookup"><span data-stu-id="f13a6-170">Name</span></span> | <span data-ttu-id="f13a6-171">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-171">Type</span></span> | <span data-ttu-id="f13a6-172">说明</span><span class="sxs-lookup"><span data-stu-id="f13a6-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="f13a6-173">字符串</span><span class="sxs-lookup"><span data-stu-id="f13a6-173">String</span></span> | <span data-ttu-id="f13a6-174">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="f13a6-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f13a6-175">Requirements</span><span class="sxs-lookup"><span data-stu-id="f13a6-175">Requirements</span></span>

|<span data-ttu-id="f13a6-176">要求</span><span class="sxs-lookup"><span data-stu-id="f13a6-176">Requirement</span></span>| <span data-ttu-id="f13a6-177">值</span><span class="sxs-lookup"><span data-stu-id="f13a6-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="f13a6-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f13a6-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f13a6-179">1.5</span><span class="sxs-lookup"><span data-stu-id="f13a6-179">1.5</span></span> |
|[<span data-ttu-id="f13a6-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f13a6-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f13a6-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f13a6-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="f13a6-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f13a6-182">SourceProperty :String</span></span>

<span data-ttu-id="f13a6-183">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="f13a6-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f13a6-184">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-184">Type</span></span>

*   <span data-ttu-id="f13a6-185">String</span><span class="sxs-lookup"><span data-stu-id="f13a6-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f13a6-186">属性：</span><span class="sxs-lookup"><span data-stu-id="f13a6-186">Properties:</span></span>

|<span data-ttu-id="f13a6-187">名称</span><span class="sxs-lookup"><span data-stu-id="f13a6-187">Name</span></span>| <span data-ttu-id="f13a6-188">类型</span><span class="sxs-lookup"><span data-stu-id="f13a6-188">Type</span></span>| <span data-ttu-id="f13a6-189">说明</span><span class="sxs-lookup"><span data-stu-id="f13a6-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f13a6-190">字符串</span><span class="sxs-lookup"><span data-stu-id="f13a6-190">String</span></span>|<span data-ttu-id="f13a6-191">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="f13a6-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f13a6-192">String</span><span class="sxs-lookup"><span data-stu-id="f13a6-192">String</span></span>|<span data-ttu-id="f13a6-193">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="f13a6-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f13a6-194">Requirements</span><span class="sxs-lookup"><span data-stu-id="f13a6-194">Requirements</span></span>

|<span data-ttu-id="f13a6-195">要求</span><span class="sxs-lookup"><span data-stu-id="f13a6-195">Requirement</span></span>| <span data-ttu-id="f13a6-196">值</span><span class="sxs-lookup"><span data-stu-id="f13a6-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="f13a6-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f13a6-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f13a6-198">1.0</span><span class="sxs-lookup"><span data-stu-id="f13a6-198">1.0</span></span>|
|[<span data-ttu-id="f13a6-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f13a6-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f13a6-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f13a6-200">Compose or Read</span></span>|
