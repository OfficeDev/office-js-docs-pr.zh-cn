---
title: Office 命名空间 - 要求集 1.5
description: ''
ms.date: 11/08/2018
localization_priority: Normal
ms.openlocfilehash: 3de90dffb3bf1baa8bc08ec512c6e89359db1349
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/13/2019
ms.locfileid: "29387644"
---
# <a name="office"></a><span data-ttu-id="71128-102">Office</span><span class="sxs-lookup"><span data-stu-id="71128-102">Office</span></span>

<span data-ttu-id="71128-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="71128-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="71128-105">要求</span><span class="sxs-lookup"><span data-stu-id="71128-105">Requirements</span></span>

|<span data-ttu-id="71128-106">要求</span><span class="sxs-lookup"><span data-stu-id="71128-106">Requirement</span></span>| <span data-ttu-id="71128-107">值</span><span class="sxs-lookup"><span data-stu-id="71128-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="71128-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71128-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71128-109">1.0</span><span class="sxs-lookup"><span data-stu-id="71128-109">1.0</span></span>|
|[<span data-ttu-id="71128-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71128-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71128-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71128-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="71128-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="71128-112">Members and methods</span></span>

| <span data-ttu-id="71128-113">成员</span><span class="sxs-lookup"><span data-stu-id="71128-113">Member</span></span> | <span data-ttu-id="71128-114">类型</span><span class="sxs-lookup"><span data-stu-id="71128-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="71128-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="71128-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="71128-116">成员</span><span class="sxs-lookup"><span data-stu-id="71128-116">Member</span></span> |
| [<span data-ttu-id="71128-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="71128-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="71128-118">成员</span><span class="sxs-lookup"><span data-stu-id="71128-118">Member</span></span> |
| [<span data-ttu-id="71128-119">EventType</span><span class="sxs-lookup"><span data-stu-id="71128-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="71128-120">成员</span><span class="sxs-lookup"><span data-stu-id="71128-120">Member</span></span> |
| [<span data-ttu-id="71128-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="71128-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="71128-122">成员</span><span class="sxs-lookup"><span data-stu-id="71128-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="71128-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="71128-123">Namespaces</span></span>

<span data-ttu-id="71128-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="71128-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="71128-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="71128-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="71128-126">成员</span><span class="sxs-lookup"><span data-stu-id="71128-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="71128-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="71128-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="71128-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="71128-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="71128-129">类型：</span><span class="sxs-lookup"><span data-stu-id="71128-129">Type:</span></span>

*   <span data-ttu-id="71128-130">字符串</span><span class="sxs-lookup"><span data-stu-id="71128-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71128-131">属性：</span><span class="sxs-lookup"><span data-stu-id="71128-131">Properties:</span></span>

|<span data-ttu-id="71128-132">名称</span><span class="sxs-lookup"><span data-stu-id="71128-132">Name</span></span>| <span data-ttu-id="71128-133">类型</span><span class="sxs-lookup"><span data-stu-id="71128-133">Type</span></span>| <span data-ttu-id="71128-134">描述</span><span class="sxs-lookup"><span data-stu-id="71128-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="71128-135">String</span><span class="sxs-lookup"><span data-stu-id="71128-135">String</span></span>|<span data-ttu-id="71128-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="71128-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="71128-137">字符串</span><span class="sxs-lookup"><span data-stu-id="71128-137">String</span></span>|<span data-ttu-id="71128-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="71128-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71128-139">要求</span><span class="sxs-lookup"><span data-stu-id="71128-139">Requirements</span></span>

|<span data-ttu-id="71128-140">要求</span><span class="sxs-lookup"><span data-stu-id="71128-140">Requirement</span></span>| <span data-ttu-id="71128-141">值</span><span class="sxs-lookup"><span data-stu-id="71128-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="71128-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71128-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71128-143">1.0</span><span class="sxs-lookup"><span data-stu-id="71128-143">1.0</span></span>|
|[<span data-ttu-id="71128-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71128-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71128-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71128-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="71128-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="71128-146">CoercionType :String</span></span>

<span data-ttu-id="71128-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="71128-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="71128-148">类型：</span><span class="sxs-lookup"><span data-stu-id="71128-148">Type:</span></span>

*   <span data-ttu-id="71128-149">字符串</span><span class="sxs-lookup"><span data-stu-id="71128-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71128-150">属性：</span><span class="sxs-lookup"><span data-stu-id="71128-150">Properties:</span></span>

|<span data-ttu-id="71128-151">名称</span><span class="sxs-lookup"><span data-stu-id="71128-151">Name</span></span>| <span data-ttu-id="71128-152">类型</span><span class="sxs-lookup"><span data-stu-id="71128-152">Type</span></span>| <span data-ttu-id="71128-153">描述</span><span class="sxs-lookup"><span data-stu-id="71128-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="71128-154">String</span><span class="sxs-lookup"><span data-stu-id="71128-154">String</span></span>|<span data-ttu-id="71128-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="71128-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="71128-156">字符串</span><span class="sxs-lookup"><span data-stu-id="71128-156">String</span></span>|<span data-ttu-id="71128-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="71128-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71128-158">要求</span><span class="sxs-lookup"><span data-stu-id="71128-158">Requirements</span></span>

|<span data-ttu-id="71128-159">要求</span><span class="sxs-lookup"><span data-stu-id="71128-159">Requirement</span></span>| <span data-ttu-id="71128-160">值</span><span class="sxs-lookup"><span data-stu-id="71128-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="71128-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71128-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71128-162">1.0</span><span class="sxs-lookup"><span data-stu-id="71128-162">1.0</span></span>|
|[<span data-ttu-id="71128-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71128-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71128-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71128-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="71128-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="71128-165">EventType :String</span></span>

<span data-ttu-id="71128-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="71128-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="71128-167">类型：</span><span class="sxs-lookup"><span data-stu-id="71128-167">Type:</span></span>

*   <span data-ttu-id="71128-168">字符串</span><span class="sxs-lookup"><span data-stu-id="71128-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71128-169">属性：</span><span class="sxs-lookup"><span data-stu-id="71128-169">Properties:</span></span>

| <span data-ttu-id="71128-170">名称</span><span class="sxs-lookup"><span data-stu-id="71128-170">Name</span></span> | <span data-ttu-id="71128-171">类型</span><span class="sxs-lookup"><span data-stu-id="71128-171">Type</span></span> | <span data-ttu-id="71128-172">描述</span><span class="sxs-lookup"><span data-stu-id="71128-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="71128-173">字符串</span><span class="sxs-lookup"><span data-stu-id="71128-173">String</span></span> | <span data-ttu-id="71128-174">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="71128-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="71128-175">要求</span><span class="sxs-lookup"><span data-stu-id="71128-175">Requirements</span></span>

|<span data-ttu-id="71128-176">要求</span><span class="sxs-lookup"><span data-stu-id="71128-176">Requirement</span></span>| <span data-ttu-id="71128-177">值</span><span class="sxs-lookup"><span data-stu-id="71128-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="71128-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71128-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71128-179">1.5</span><span class="sxs-lookup"><span data-stu-id="71128-179">1.5</span></span> |
|[<span data-ttu-id="71128-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71128-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71128-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71128-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="71128-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="71128-182">SourceProperty :String</span></span>

<span data-ttu-id="71128-183">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="71128-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="71128-184">类型：</span><span class="sxs-lookup"><span data-stu-id="71128-184">Type:</span></span>

*   <span data-ttu-id="71128-185">String</span><span class="sxs-lookup"><span data-stu-id="71128-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71128-186">属性：</span><span class="sxs-lookup"><span data-stu-id="71128-186">Properties:</span></span>

|<span data-ttu-id="71128-187">名称</span><span class="sxs-lookup"><span data-stu-id="71128-187">Name</span></span>| <span data-ttu-id="71128-188">类型</span><span class="sxs-lookup"><span data-stu-id="71128-188">Type</span></span>| <span data-ttu-id="71128-189">说明</span><span class="sxs-lookup"><span data-stu-id="71128-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="71128-190">字符串</span><span class="sxs-lookup"><span data-stu-id="71128-190">String</span></span>|<span data-ttu-id="71128-191">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="71128-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="71128-192">String</span><span class="sxs-lookup"><span data-stu-id="71128-192">String</span></span>|<span data-ttu-id="71128-193">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="71128-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71128-194">要求</span><span class="sxs-lookup"><span data-stu-id="71128-194">Requirements</span></span>

|<span data-ttu-id="71128-195">要求</span><span class="sxs-lookup"><span data-stu-id="71128-195">Requirement</span></span>| <span data-ttu-id="71128-196">值</span><span class="sxs-lookup"><span data-stu-id="71128-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="71128-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="71128-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71128-198">1.0</span><span class="sxs-lookup"><span data-stu-id="71128-198">1.0</span></span>|
|[<span data-ttu-id="71128-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="71128-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71128-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="71128-200">Compose or read</span></span>|
