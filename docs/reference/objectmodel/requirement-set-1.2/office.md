---
title: Office 命名空间-要求集1。2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 126e9392980510656bf9da8cf760b616b623d153
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268703"
---
# <a name="office"></a><span data-ttu-id="d21f8-102">Office</span><span class="sxs-lookup"><span data-stu-id="d21f8-102">Office</span></span>

<span data-ttu-id="d21f8-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="d21f8-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d21f8-105">要求</span><span class="sxs-lookup"><span data-stu-id="d21f8-105">Requirements</span></span>

|<span data-ttu-id="d21f8-106">要求</span><span class="sxs-lookup"><span data-stu-id="d21f8-106">Requirement</span></span>| <span data-ttu-id="d21f8-107">值</span><span class="sxs-lookup"><span data-stu-id="d21f8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d21f8-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d21f8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d21f8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d21f8-109">1.0</span></span>|
|[<span data-ttu-id="d21f8-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d21f8-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d21f8-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d21f8-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d21f8-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="d21f8-112">Members and methods</span></span>

| <span data-ttu-id="d21f8-113">成员</span><span class="sxs-lookup"><span data-stu-id="d21f8-113">Member</span></span> | <span data-ttu-id="d21f8-114">类型</span><span class="sxs-lookup"><span data-stu-id="d21f8-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d21f8-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d21f8-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d21f8-116">Member</span><span class="sxs-lookup"><span data-stu-id="d21f8-116">Member</span></span> |
| [<span data-ttu-id="d21f8-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d21f8-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d21f8-118">Member</span><span class="sxs-lookup"><span data-stu-id="d21f8-118">Member</span></span> |
| [<span data-ttu-id="d21f8-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d21f8-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d21f8-120">成员</span><span class="sxs-lookup"><span data-stu-id="d21f8-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d21f8-121">命名空间</span><span class="sxs-lookup"><span data-stu-id="d21f8-121">Namespaces</span></span>

<span data-ttu-id="d21f8-122">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="d21f8-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d21f8-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="d21f8-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d21f8-124">成员</span><span class="sxs-lookup"><span data-stu-id="d21f8-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d21f8-125">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="d21f8-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="d21f8-126">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="d21f8-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d21f8-127">类型</span><span class="sxs-lookup"><span data-stu-id="d21f8-127">Type</span></span>

*   <span data-ttu-id="d21f8-128">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d21f8-129">属性：</span><span class="sxs-lookup"><span data-stu-id="d21f8-129">Properties:</span></span>

|<span data-ttu-id="d21f8-130">名称</span><span class="sxs-lookup"><span data-stu-id="d21f8-130">Name</span></span>| <span data-ttu-id="d21f8-131">类型</span><span class="sxs-lookup"><span data-stu-id="d21f8-131">Type</span></span>| <span data-ttu-id="d21f8-132">说明</span><span class="sxs-lookup"><span data-stu-id="d21f8-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d21f8-133">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-133">String</span></span>|<span data-ttu-id="d21f8-134">调用成功。</span><span class="sxs-lookup"><span data-stu-id="d21f8-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d21f8-135">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-135">String</span></span>|<span data-ttu-id="d21f8-136">调用失败。</span><span class="sxs-lookup"><span data-stu-id="d21f8-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d21f8-137">要求</span><span class="sxs-lookup"><span data-stu-id="d21f8-137">Requirements</span></span>

|<span data-ttu-id="d21f8-138">要求</span><span class="sxs-lookup"><span data-stu-id="d21f8-138">Requirement</span></span>| <span data-ttu-id="d21f8-139">值</span><span class="sxs-lookup"><span data-stu-id="d21f8-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="d21f8-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d21f8-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d21f8-141">1.0</span><span class="sxs-lookup"><span data-stu-id="d21f8-141">1.0</span></span>|
|[<span data-ttu-id="d21f8-142">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d21f8-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d21f8-143">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d21f8-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="d21f8-144">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="d21f8-144">CoercionType: String</span></span>

<span data-ttu-id="d21f8-145">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="d21f8-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d21f8-146">类型</span><span class="sxs-lookup"><span data-stu-id="d21f8-146">Type</span></span>

*   <span data-ttu-id="d21f8-147">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d21f8-148">属性：</span><span class="sxs-lookup"><span data-stu-id="d21f8-148">Properties:</span></span>

|<span data-ttu-id="d21f8-149">名称</span><span class="sxs-lookup"><span data-stu-id="d21f8-149">Name</span></span>| <span data-ttu-id="d21f8-150">类型</span><span class="sxs-lookup"><span data-stu-id="d21f8-150">Type</span></span>| <span data-ttu-id="d21f8-151">说明</span><span class="sxs-lookup"><span data-stu-id="d21f8-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d21f8-152">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-152">String</span></span>|<span data-ttu-id="d21f8-153">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d21f8-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d21f8-154">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-154">String</span></span>|<span data-ttu-id="d21f8-155">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d21f8-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d21f8-156">要求</span><span class="sxs-lookup"><span data-stu-id="d21f8-156">Requirements</span></span>

|<span data-ttu-id="d21f8-157">要求</span><span class="sxs-lookup"><span data-stu-id="d21f8-157">Requirement</span></span>| <span data-ttu-id="d21f8-158">值</span><span class="sxs-lookup"><span data-stu-id="d21f8-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="d21f8-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d21f8-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d21f8-160">1.0</span><span class="sxs-lookup"><span data-stu-id="d21f8-160">1.0</span></span>|
|[<span data-ttu-id="d21f8-161">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d21f8-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d21f8-162">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d21f8-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="d21f8-163">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="d21f8-163">SourceProperty: String</span></span>

<span data-ttu-id="d21f8-164">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="d21f8-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d21f8-165">类型</span><span class="sxs-lookup"><span data-stu-id="d21f8-165">Type</span></span>

*   <span data-ttu-id="d21f8-166">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d21f8-167">属性：</span><span class="sxs-lookup"><span data-stu-id="d21f8-167">Properties:</span></span>

|<span data-ttu-id="d21f8-168">名称</span><span class="sxs-lookup"><span data-stu-id="d21f8-168">Name</span></span>| <span data-ttu-id="d21f8-169">类型</span><span class="sxs-lookup"><span data-stu-id="d21f8-169">Type</span></span>| <span data-ttu-id="d21f8-170">说明</span><span class="sxs-lookup"><span data-stu-id="d21f8-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d21f8-171">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-171">String</span></span>|<span data-ttu-id="d21f8-172">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="d21f8-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d21f8-173">String</span><span class="sxs-lookup"><span data-stu-id="d21f8-173">String</span></span>|<span data-ttu-id="d21f8-174">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="d21f8-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d21f8-175">要求</span><span class="sxs-lookup"><span data-stu-id="d21f8-175">Requirements</span></span>

|<span data-ttu-id="d21f8-176">要求</span><span class="sxs-lookup"><span data-stu-id="d21f8-176">Requirement</span></span>| <span data-ttu-id="d21f8-177">值</span><span class="sxs-lookup"><span data-stu-id="d21f8-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="d21f8-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d21f8-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d21f8-179">1.0</span><span class="sxs-lookup"><span data-stu-id="d21f8-179">1.0</span></span>|
|[<span data-ttu-id="d21f8-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d21f8-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d21f8-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d21f8-181">Compose or Read</span></span>|
