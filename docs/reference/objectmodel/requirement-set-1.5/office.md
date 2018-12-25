---
title: Office 命名空间 - 要求集 1.5
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 46b21df77456d2392fbc543e45513246a4ad9a10
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433647"
---
# <a name="office"></a><span data-ttu-id="6496a-102">Office</span><span class="sxs-lookup"><span data-stu-id="6496a-102">Office</span></span>

<span data-ttu-id="6496a-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="6496a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6496a-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="6496a-105">Requirements</span></span>

|<span data-ttu-id="6496a-106">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-106">Requirement</span></span>| <span data-ttu-id="6496a-107">值</span><span class="sxs-lookup"><span data-stu-id="6496a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6496a-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6496a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6496a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6496a-109">1.0</span></span>|
|[<span data-ttu-id="6496a-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6496a-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6496a-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6496a-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6496a-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="6496a-112">Members and methods</span></span>

| <span data-ttu-id="6496a-113">成员</span><span class="sxs-lookup"><span data-stu-id="6496a-113">Member</span></span> | <span data-ttu-id="6496a-114">类型</span><span class="sxs-lookup"><span data-stu-id="6496a-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6496a-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6496a-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6496a-116">成员</span><span class="sxs-lookup"><span data-stu-id="6496a-116">Member</span></span> |
| [<span data-ttu-id="6496a-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6496a-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6496a-118">成员</span><span class="sxs-lookup"><span data-stu-id="6496a-118">Member</span></span> |
| [<span data-ttu-id="6496a-119">EventType</span><span class="sxs-lookup"><span data-stu-id="6496a-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6496a-120">成员</span><span class="sxs-lookup"><span data-stu-id="6496a-120">Member</span></span> |
| [<span data-ttu-id="6496a-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6496a-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6496a-122">成员</span><span class="sxs-lookup"><span data-stu-id="6496a-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="6496a-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="6496a-123">Namespaces</span></span>

<span data-ttu-id="6496a-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="6496a-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="6496a-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="6496a-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="6496a-126">成员</span><span class="sxs-lookup"><span data-stu-id="6496a-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="6496a-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="6496a-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="6496a-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="6496a-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6496a-129">类型：</span><span class="sxs-lookup"><span data-stu-id="6496a-129">Type:</span></span>

*   <span data-ttu-id="6496a-130">字符串</span><span class="sxs-lookup"><span data-stu-id="6496a-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6496a-131">属性：</span><span class="sxs-lookup"><span data-stu-id="6496a-131">Properties:</span></span>

|<span data-ttu-id="6496a-132">名称</span><span class="sxs-lookup"><span data-stu-id="6496a-132">Name</span></span>| <span data-ttu-id="6496a-133">类型</span><span class="sxs-lookup"><span data-stu-id="6496a-133">Type</span></span>| <span data-ttu-id="6496a-134">描述</span><span class="sxs-lookup"><span data-stu-id="6496a-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6496a-135">String</span><span class="sxs-lookup"><span data-stu-id="6496a-135">String</span></span>|<span data-ttu-id="6496a-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="6496a-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6496a-137">字符串</span><span class="sxs-lookup"><span data-stu-id="6496a-137">String</span></span>|<span data-ttu-id="6496a-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="6496a-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6496a-139">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-139">Requirements</span></span>

|<span data-ttu-id="6496a-140">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-140">Requirement</span></span>| <span data-ttu-id="6496a-141">值</span><span class="sxs-lookup"><span data-stu-id="6496a-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="6496a-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6496a-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6496a-143">1.0</span><span class="sxs-lookup"><span data-stu-id="6496a-143">1.0</span></span>|
|[<span data-ttu-id="6496a-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6496a-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6496a-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6496a-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="6496a-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="6496a-146">CoercionType :String</span></span>

<span data-ttu-id="6496a-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="6496a-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6496a-148">类型：</span><span class="sxs-lookup"><span data-stu-id="6496a-148">Type:</span></span>

*   <span data-ttu-id="6496a-149">字符串</span><span class="sxs-lookup"><span data-stu-id="6496a-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6496a-150">属性：</span><span class="sxs-lookup"><span data-stu-id="6496a-150">Properties:</span></span>

|<span data-ttu-id="6496a-151">名称</span><span class="sxs-lookup"><span data-stu-id="6496a-151">Name</span></span>| <span data-ttu-id="6496a-152">类型</span><span class="sxs-lookup"><span data-stu-id="6496a-152">Type</span></span>| <span data-ttu-id="6496a-153">描述</span><span class="sxs-lookup"><span data-stu-id="6496a-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6496a-154">String</span><span class="sxs-lookup"><span data-stu-id="6496a-154">String</span></span>|<span data-ttu-id="6496a-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="6496a-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6496a-156">字符串</span><span class="sxs-lookup"><span data-stu-id="6496a-156">String</span></span>|<span data-ttu-id="6496a-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="6496a-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6496a-158">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-158">Requirements</span></span>

|<span data-ttu-id="6496a-159">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-159">Requirement</span></span>| <span data-ttu-id="6496a-160">值</span><span class="sxs-lookup"><span data-stu-id="6496a-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="6496a-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6496a-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6496a-162">1.0</span><span class="sxs-lookup"><span data-stu-id="6496a-162">1.0</span></span>|
|[<span data-ttu-id="6496a-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6496a-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6496a-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6496a-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="6496a-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="6496a-165">EventType :String</span></span>

<span data-ttu-id="6496a-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="6496a-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6496a-167">类型：</span><span class="sxs-lookup"><span data-stu-id="6496a-167">Type:</span></span>

*   <span data-ttu-id="6496a-168">字符串</span><span class="sxs-lookup"><span data-stu-id="6496a-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6496a-169">属性：</span><span class="sxs-lookup"><span data-stu-id="6496a-169">Properties:</span></span>

| <span data-ttu-id="6496a-170">名称</span><span class="sxs-lookup"><span data-stu-id="6496a-170">Name</span></span> | <span data-ttu-id="6496a-171">类型</span><span class="sxs-lookup"><span data-stu-id="6496a-171">Type</span></span> | <span data-ttu-id="6496a-172">描述</span><span class="sxs-lookup"><span data-stu-id="6496a-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="6496a-173">字符串</span><span class="sxs-lookup"><span data-stu-id="6496a-173">String</span></span> | <span data-ttu-id="6496a-174">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="6496a-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6496a-175">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-175">Requirements</span></span>

|<span data-ttu-id="6496a-176">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-176">Requirement</span></span>| <span data-ttu-id="6496a-177">值</span><span class="sxs-lookup"><span data-stu-id="6496a-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="6496a-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6496a-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6496a-179">1.5</span><span class="sxs-lookup"><span data-stu-id="6496a-179">1.5</span></span> |
|[<span data-ttu-id="6496a-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6496a-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6496a-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6496a-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="6496a-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="6496a-182">SourceProperty :String</span></span>

<span data-ttu-id="6496a-183">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="6496a-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6496a-184">类型：</span><span class="sxs-lookup"><span data-stu-id="6496a-184">Type:</span></span>

*   <span data-ttu-id="6496a-185">字符串</span><span class="sxs-lookup"><span data-stu-id="6496a-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6496a-186">属性：</span><span class="sxs-lookup"><span data-stu-id="6496a-186">Properties:</span></span>

|<span data-ttu-id="6496a-187">名称</span><span class="sxs-lookup"><span data-stu-id="6496a-187">Name</span></span>| <span data-ttu-id="6496a-188">类型</span><span class="sxs-lookup"><span data-stu-id="6496a-188">Type</span></span>| <span data-ttu-id="6496a-189">描述</span><span class="sxs-lookup"><span data-stu-id="6496a-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6496a-190">字符串</span><span class="sxs-lookup"><span data-stu-id="6496a-190">String</span></span>|<span data-ttu-id="6496a-191">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="6496a-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6496a-192">String</span><span class="sxs-lookup"><span data-stu-id="6496a-192">String</span></span>|<span data-ttu-id="6496a-193">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="6496a-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6496a-194">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-194">Requirements</span></span>

|<span data-ttu-id="6496a-195">要求</span><span class="sxs-lookup"><span data-stu-id="6496a-195">Requirement</span></span>| <span data-ttu-id="6496a-196">值</span><span class="sxs-lookup"><span data-stu-id="6496a-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="6496a-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6496a-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6496a-198">1.0</span><span class="sxs-lookup"><span data-stu-id="6496a-198">1.0</span></span>|
|[<span data-ttu-id="6496a-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6496a-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6496a-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6496a-200">Compose or read</span></span>|