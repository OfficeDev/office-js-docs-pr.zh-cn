---
title: Office 命名空间-要求集1。6
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 84e8fa49e1d4dce4239525badafaa051325bb3ec
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395636"
---
# <a name="office"></a><span data-ttu-id="89f66-102">Office</span><span class="sxs-lookup"><span data-stu-id="89f66-102">Office</span></span>

<span data-ttu-id="89f66-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="89f66-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="89f66-105">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-105">Requirements</span></span>

|<span data-ttu-id="89f66-106">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-106">Requirement</span></span>| <span data-ttu-id="89f66-107">值</span><span class="sxs-lookup"><span data-stu-id="89f66-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="89f66-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="89f66-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89f66-109">1.0</span><span class="sxs-lookup"><span data-stu-id="89f66-109">1.0</span></span>|
|[<span data-ttu-id="89f66-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="89f66-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89f66-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="89f66-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="89f66-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="89f66-112">Members and methods</span></span>

| <span data-ttu-id="89f66-113">成员</span><span class="sxs-lookup"><span data-stu-id="89f66-113">Member</span></span> | <span data-ttu-id="89f66-114">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="89f66-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="89f66-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="89f66-116">Member</span><span class="sxs-lookup"><span data-stu-id="89f66-116">Member</span></span> |
| [<span data-ttu-id="89f66-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="89f66-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="89f66-118">Member</span><span class="sxs-lookup"><span data-stu-id="89f66-118">Member</span></span> |
| [<span data-ttu-id="89f66-119">EventType</span><span class="sxs-lookup"><span data-stu-id="89f66-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="89f66-120">Member</span><span class="sxs-lookup"><span data-stu-id="89f66-120">Member</span></span> |
| [<span data-ttu-id="89f66-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="89f66-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="89f66-122">成员</span><span class="sxs-lookup"><span data-stu-id="89f66-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="89f66-123">命名空间</span><span class="sxs-lookup"><span data-stu-id="89f66-123">Namespaces</span></span>

<span data-ttu-id="89f66-124">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="89f66-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="89f66-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): 包含多个`ItemType`枚举, 例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="89f66-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="89f66-126">Members</span><span class="sxs-lookup"><span data-stu-id="89f66-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="89f66-127">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="89f66-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="89f66-128">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="89f66-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="89f66-129">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-129">Type</span></span>

*   <span data-ttu-id="89f66-130">String</span><span class="sxs-lookup"><span data-stu-id="89f66-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="89f66-131">属性：</span><span class="sxs-lookup"><span data-stu-id="89f66-131">Properties:</span></span>

|<span data-ttu-id="89f66-132">名称</span><span class="sxs-lookup"><span data-stu-id="89f66-132">Name</span></span>| <span data-ttu-id="89f66-133">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-133">Type</span></span>| <span data-ttu-id="89f66-134">说明</span><span class="sxs-lookup"><span data-stu-id="89f66-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="89f66-135">String</span><span class="sxs-lookup"><span data-stu-id="89f66-135">String</span></span>|<span data-ttu-id="89f66-136">调用成功。</span><span class="sxs-lookup"><span data-stu-id="89f66-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="89f66-137">String</span><span class="sxs-lookup"><span data-stu-id="89f66-137">String</span></span>|<span data-ttu-id="89f66-138">调用失败。</span><span class="sxs-lookup"><span data-stu-id="89f66-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="89f66-139">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-139">Requirements</span></span>

|<span data-ttu-id="89f66-140">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-140">Requirement</span></span>| <span data-ttu-id="89f66-141">值</span><span class="sxs-lookup"><span data-stu-id="89f66-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="89f66-142">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="89f66-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89f66-143">1.0</span><span class="sxs-lookup"><span data-stu-id="89f66-143">1.0</span></span>|
|[<span data-ttu-id="89f66-144">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="89f66-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89f66-145">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="89f66-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="89f66-146">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="89f66-146">CoercionType: String</span></span>

<span data-ttu-id="89f66-147">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="89f66-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="89f66-148">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-148">Type</span></span>

*   <span data-ttu-id="89f66-149">String</span><span class="sxs-lookup"><span data-stu-id="89f66-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="89f66-150">属性：</span><span class="sxs-lookup"><span data-stu-id="89f66-150">Properties:</span></span>

|<span data-ttu-id="89f66-151">名称</span><span class="sxs-lookup"><span data-stu-id="89f66-151">Name</span></span>| <span data-ttu-id="89f66-152">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-152">Type</span></span>| <span data-ttu-id="89f66-153">说明</span><span class="sxs-lookup"><span data-stu-id="89f66-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="89f66-154">String</span><span class="sxs-lookup"><span data-stu-id="89f66-154">String</span></span>|<span data-ttu-id="89f66-155">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="89f66-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="89f66-156">String</span><span class="sxs-lookup"><span data-stu-id="89f66-156">String</span></span>|<span data-ttu-id="89f66-157">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="89f66-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="89f66-158">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-158">Requirements</span></span>

|<span data-ttu-id="89f66-159">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-159">Requirement</span></span>| <span data-ttu-id="89f66-160">值</span><span class="sxs-lookup"><span data-stu-id="89f66-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="89f66-161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="89f66-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89f66-162">1.0</span><span class="sxs-lookup"><span data-stu-id="89f66-162">1.0</span></span>|
|[<span data-ttu-id="89f66-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="89f66-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89f66-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="89f66-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="89f66-165">事件类型: String</span><span class="sxs-lookup"><span data-stu-id="89f66-165">EventType: String</span></span>

<span data-ttu-id="89f66-166">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="89f66-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="89f66-167">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-167">Type</span></span>

*   <span data-ttu-id="89f66-168">String</span><span class="sxs-lookup"><span data-stu-id="89f66-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="89f66-169">属性：</span><span class="sxs-lookup"><span data-stu-id="89f66-169">Properties:</span></span>

| <span data-ttu-id="89f66-170">名称</span><span class="sxs-lookup"><span data-stu-id="89f66-170">Name</span></span> | <span data-ttu-id="89f66-171">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-171">Type</span></span> | <span data-ttu-id="89f66-172">说明</span><span class="sxs-lookup"><span data-stu-id="89f66-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="89f66-173">String</span><span class="sxs-lookup"><span data-stu-id="89f66-173">String</span></span> | <span data-ttu-id="89f66-174">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="89f66-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="89f66-175">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-175">Requirements</span></span>

|<span data-ttu-id="89f66-176">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-176">Requirement</span></span>| <span data-ttu-id="89f66-177">值</span><span class="sxs-lookup"><span data-stu-id="89f66-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="89f66-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="89f66-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89f66-179">1.5</span><span class="sxs-lookup"><span data-stu-id="89f66-179">1.5</span></span> |
|[<span data-ttu-id="89f66-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="89f66-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89f66-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="89f66-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="89f66-182">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="89f66-182">SourceProperty: String</span></span>

<span data-ttu-id="89f66-183">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="89f66-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="89f66-184">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-184">Type</span></span>

*   <span data-ttu-id="89f66-185">String</span><span class="sxs-lookup"><span data-stu-id="89f66-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="89f66-186">属性：</span><span class="sxs-lookup"><span data-stu-id="89f66-186">Properties:</span></span>

|<span data-ttu-id="89f66-187">名称</span><span class="sxs-lookup"><span data-stu-id="89f66-187">Name</span></span>| <span data-ttu-id="89f66-188">类型</span><span class="sxs-lookup"><span data-stu-id="89f66-188">Type</span></span>| <span data-ttu-id="89f66-189">说明</span><span class="sxs-lookup"><span data-stu-id="89f66-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="89f66-190">String</span><span class="sxs-lookup"><span data-stu-id="89f66-190">String</span></span>|<span data-ttu-id="89f66-191">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="89f66-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="89f66-192">String</span><span class="sxs-lookup"><span data-stu-id="89f66-192">String</span></span>|<span data-ttu-id="89f66-193">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="89f66-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="89f66-194">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-194">Requirements</span></span>

|<span data-ttu-id="89f66-195">要求</span><span class="sxs-lookup"><span data-stu-id="89f66-195">Requirement</span></span>| <span data-ttu-id="89f66-196">值</span><span class="sxs-lookup"><span data-stu-id="89f66-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="89f66-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="89f66-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89f66-198">1.0</span><span class="sxs-lookup"><span data-stu-id="89f66-198">1.0</span></span>|
|[<span data-ttu-id="89f66-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="89f66-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89f66-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="89f66-200">Compose or Read</span></span>|
