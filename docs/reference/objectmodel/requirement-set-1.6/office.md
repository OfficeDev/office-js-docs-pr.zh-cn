---
title: Office 命名空间 - 要求集 1.6
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 73411efee9dcfffa5f9f0fa9de85dafc31a4173a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432282"
---
# <a name="office"></a><span data-ttu-id="0c182-102">Office</span><span class="sxs-lookup"><span data-stu-id="0c182-102">Office</span></span>

<span data-ttu-id="0c182-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="0c182-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0c182-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="0c182-105">Requirements</span></span>

|<span data-ttu-id="0c182-106">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-106">Requirement</span></span>| <span data-ttu-id="0c182-107">值</span><span class="sxs-lookup"><span data-stu-id="0c182-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c182-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0c182-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c182-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0c182-109">1.0</span></span>|
|[<span data-ttu-id="0c182-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0c182-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c182-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0c182-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0c182-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="0c182-112">Members and methods</span></span>

| <span data-ttu-id="0c182-113">成员</span><span class="sxs-lookup"><span data-stu-id="0c182-113">Member</span></span> | <span data-ttu-id="0c182-114">类型</span><span class="sxs-lookup"><span data-stu-id="0c182-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0c182-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0c182-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0c182-116">成员</span><span class="sxs-lookup"><span data-stu-id="0c182-116">Member</span></span> |
| [<span data-ttu-id="0c182-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0c182-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0c182-118">成员</span><span class="sxs-lookup"><span data-stu-id="0c182-118">Member</span></span> |
| [<span data-ttu-id="0c182-119">EventType</span><span class="sxs-lookup"><span data-stu-id="0c182-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="0c182-120">成员</span><span class="sxs-lookup"><span data-stu-id="0c182-120">Member</span></span> |
| [<span data-ttu-id="0c182-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0c182-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0c182-122">成员</span><span class="sxs-lookup"><span data-stu-id="0c182-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0c182-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="0c182-123">Namespaces</span></span>

<span data-ttu-id="0c182-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="0c182-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="0c182-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="0c182-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="0c182-126">成员</span><span class="sxs-lookup"><span data-stu-id="0c182-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="0c182-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="0c182-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="0c182-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="0c182-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0c182-129">类型：</span><span class="sxs-lookup"><span data-stu-id="0c182-129">Type:</span></span>

*   <span data-ttu-id="0c182-130">字符串</span><span class="sxs-lookup"><span data-stu-id="0c182-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0c182-131">属性：</span><span class="sxs-lookup"><span data-stu-id="0c182-131">Properties:</span></span>

|<span data-ttu-id="0c182-132">名称</span><span class="sxs-lookup"><span data-stu-id="0c182-132">Name</span></span>| <span data-ttu-id="0c182-133">类型</span><span class="sxs-lookup"><span data-stu-id="0c182-133">Type</span></span>| <span data-ttu-id="0c182-134">描述</span><span class="sxs-lookup"><span data-stu-id="0c182-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0c182-135">String</span><span class="sxs-lookup"><span data-stu-id="0c182-135">String</span></span>|<span data-ttu-id="0c182-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="0c182-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0c182-137">字符串</span><span class="sxs-lookup"><span data-stu-id="0c182-137">String</span></span>|<span data-ttu-id="0c182-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="0c182-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c182-139">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-139">Requirements</span></span>

|<span data-ttu-id="0c182-140">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-140">Requirement</span></span>| <span data-ttu-id="0c182-141">值</span><span class="sxs-lookup"><span data-stu-id="0c182-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c182-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0c182-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c182-143">1.0</span><span class="sxs-lookup"><span data-stu-id="0c182-143">1.0</span></span>|
|[<span data-ttu-id="0c182-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0c182-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c182-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0c182-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="0c182-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="0c182-146">CoercionType :String</span></span>

<span data-ttu-id="0c182-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="0c182-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0c182-148">类型：</span><span class="sxs-lookup"><span data-stu-id="0c182-148">Type:</span></span>

*   <span data-ttu-id="0c182-149">字符串</span><span class="sxs-lookup"><span data-stu-id="0c182-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0c182-150">属性：</span><span class="sxs-lookup"><span data-stu-id="0c182-150">Properties:</span></span>

|<span data-ttu-id="0c182-151">名称</span><span class="sxs-lookup"><span data-stu-id="0c182-151">Name</span></span>| <span data-ttu-id="0c182-152">类型</span><span class="sxs-lookup"><span data-stu-id="0c182-152">Type</span></span>| <span data-ttu-id="0c182-153">描述</span><span class="sxs-lookup"><span data-stu-id="0c182-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0c182-154">String</span><span class="sxs-lookup"><span data-stu-id="0c182-154">String</span></span>|<span data-ttu-id="0c182-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="0c182-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0c182-156">字符串</span><span class="sxs-lookup"><span data-stu-id="0c182-156">String</span></span>|<span data-ttu-id="0c182-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="0c182-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c182-158">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-158">Requirements</span></span>

|<span data-ttu-id="0c182-159">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-159">Requirement</span></span>| <span data-ttu-id="0c182-160">值</span><span class="sxs-lookup"><span data-stu-id="0c182-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c182-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0c182-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c182-162">1.0</span><span class="sxs-lookup"><span data-stu-id="0c182-162">1.0</span></span>|
|[<span data-ttu-id="0c182-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0c182-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c182-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0c182-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="0c182-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="0c182-165">EventType :String</span></span>

<span data-ttu-id="0c182-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="0c182-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="0c182-167">类型：</span><span class="sxs-lookup"><span data-stu-id="0c182-167">Type:</span></span>

*   <span data-ttu-id="0c182-168">字符串</span><span class="sxs-lookup"><span data-stu-id="0c182-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0c182-169">属性：</span><span class="sxs-lookup"><span data-stu-id="0c182-169">Properties:</span></span>

| <span data-ttu-id="0c182-170">名称</span><span class="sxs-lookup"><span data-stu-id="0c182-170">Name</span></span> | <span data-ttu-id="0c182-171">类型</span><span class="sxs-lookup"><span data-stu-id="0c182-171">Type</span></span> | <span data-ttu-id="0c182-172">描述</span><span class="sxs-lookup"><span data-stu-id="0c182-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="0c182-173">字符串</span><span class="sxs-lookup"><span data-stu-id="0c182-173">String</span></span> | <span data-ttu-id="0c182-174">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="0c182-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0c182-175">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-175">Requirements</span></span>

|<span data-ttu-id="0c182-176">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-176">Requirement</span></span>| <span data-ttu-id="0c182-177">值</span><span class="sxs-lookup"><span data-stu-id="0c182-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c182-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0c182-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c182-179">1.5</span><span class="sxs-lookup"><span data-stu-id="0c182-179">1.5</span></span> |
|[<span data-ttu-id="0c182-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0c182-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c182-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0c182-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="0c182-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="0c182-182">SourceProperty :String</span></span>

<span data-ttu-id="0c182-183">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="0c182-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0c182-184">类型：</span><span class="sxs-lookup"><span data-stu-id="0c182-184">Type:</span></span>

*   <span data-ttu-id="0c182-185">字符串</span><span class="sxs-lookup"><span data-stu-id="0c182-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0c182-186">属性：</span><span class="sxs-lookup"><span data-stu-id="0c182-186">Properties:</span></span>

|<span data-ttu-id="0c182-187">名称</span><span class="sxs-lookup"><span data-stu-id="0c182-187">Name</span></span>| <span data-ttu-id="0c182-188">类型</span><span class="sxs-lookup"><span data-stu-id="0c182-188">Type</span></span>| <span data-ttu-id="0c182-189">描述</span><span class="sxs-lookup"><span data-stu-id="0c182-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0c182-190">字符串</span><span class="sxs-lookup"><span data-stu-id="0c182-190">String</span></span>|<span data-ttu-id="0c182-191">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="0c182-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0c182-192">String</span><span class="sxs-lookup"><span data-stu-id="0c182-192">String</span></span>|<span data-ttu-id="0c182-193">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="0c182-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c182-194">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-194">Requirements</span></span>

|<span data-ttu-id="0c182-195">要求</span><span class="sxs-lookup"><span data-stu-id="0c182-195">Requirement</span></span>| <span data-ttu-id="0c182-196">值</span><span class="sxs-lookup"><span data-stu-id="0c182-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c182-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0c182-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c182-198">1.0</span><span class="sxs-lookup"><span data-stu-id="0c182-198">1.0</span></span>|
|[<span data-ttu-id="0c182-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0c182-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c182-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0c182-200">Compose or read</span></span>|