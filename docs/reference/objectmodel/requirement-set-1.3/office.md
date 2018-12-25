---
title: Office 命名空间 - 要求集 1.3
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: c269f21b98e7f87d6f064f6c8ea0c439916f7caf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432744"
---
# <a name="office"></a><span data-ttu-id="86de7-102">Office</span><span class="sxs-lookup"><span data-stu-id="86de7-102">Office</span></span>

<span data-ttu-id="86de7-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="86de7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="86de7-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="86de7-105">Requirements</span></span>

|<span data-ttu-id="86de7-106">要求</span><span class="sxs-lookup"><span data-stu-id="86de7-106">Requirement</span></span>| <span data-ttu-id="86de7-107">值</span><span class="sxs-lookup"><span data-stu-id="86de7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="86de7-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="86de7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86de7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="86de7-109">1.0</span></span>|
|[<span data-ttu-id="86de7-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="86de7-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="86de7-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="86de7-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="86de7-112">命名空间</span><span class="sxs-lookup"><span data-stu-id="86de7-112">Namespaces</span></span>

<span data-ttu-id="86de7-113">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="86de7-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="86de7-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="86de7-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="86de7-115">成员</span><span class="sxs-lookup"><span data-stu-id="86de7-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="86de7-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="86de7-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="86de7-117">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="86de7-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="86de7-118">类型：</span><span class="sxs-lookup"><span data-stu-id="86de7-118">Type:</span></span>

*   <span data-ttu-id="86de7-119">字符串</span><span class="sxs-lookup"><span data-stu-id="86de7-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="86de7-120">属性：</span><span class="sxs-lookup"><span data-stu-id="86de7-120">Properties:</span></span>

|<span data-ttu-id="86de7-121">名称</span><span class="sxs-lookup"><span data-stu-id="86de7-121">Name</span></span>| <span data-ttu-id="86de7-122">类型</span><span class="sxs-lookup"><span data-stu-id="86de7-122">Type</span></span>| <span data-ttu-id="86de7-123">描述</span><span class="sxs-lookup"><span data-stu-id="86de7-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="86de7-124">String</span><span class="sxs-lookup"><span data-stu-id="86de7-124">String</span></span>|<span data-ttu-id="86de7-125">调用成功。</span><span class="sxs-lookup"><span data-stu-id="86de7-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="86de7-126">字符串</span><span class="sxs-lookup"><span data-stu-id="86de7-126">String</span></span>|<span data-ttu-id="86de7-127">调用失败。</span><span class="sxs-lookup"><span data-stu-id="86de7-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86de7-128">要求</span><span class="sxs-lookup"><span data-stu-id="86de7-128">Requirements</span></span>

|<span data-ttu-id="86de7-129">要求</span><span class="sxs-lookup"><span data-stu-id="86de7-129">Requirement</span></span>| <span data-ttu-id="86de7-130">值</span><span class="sxs-lookup"><span data-stu-id="86de7-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="86de7-131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="86de7-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86de7-132">1.0</span><span class="sxs-lookup"><span data-stu-id="86de7-132">1.0</span></span>|
|[<span data-ttu-id="86de7-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="86de7-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="86de7-134">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="86de7-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="86de7-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="86de7-135">CoercionType :String</span></span>

<span data-ttu-id="86de7-136">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="86de7-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="86de7-137">类型：</span><span class="sxs-lookup"><span data-stu-id="86de7-137">Type:</span></span>

*   <span data-ttu-id="86de7-138">字符串</span><span class="sxs-lookup"><span data-stu-id="86de7-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="86de7-139">属性：</span><span class="sxs-lookup"><span data-stu-id="86de7-139">Properties:</span></span>

|<span data-ttu-id="86de7-140">名称</span><span class="sxs-lookup"><span data-stu-id="86de7-140">Name</span></span>| <span data-ttu-id="86de7-141">类型</span><span class="sxs-lookup"><span data-stu-id="86de7-141">Type</span></span>| <span data-ttu-id="86de7-142">描述</span><span class="sxs-lookup"><span data-stu-id="86de7-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="86de7-143">String</span><span class="sxs-lookup"><span data-stu-id="86de7-143">String</span></span>|<span data-ttu-id="86de7-144">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="86de7-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="86de7-145">字符串</span><span class="sxs-lookup"><span data-stu-id="86de7-145">String</span></span>|<span data-ttu-id="86de7-146">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="86de7-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86de7-147">要求</span><span class="sxs-lookup"><span data-stu-id="86de7-147">Requirements</span></span>

|<span data-ttu-id="86de7-148">要求</span><span class="sxs-lookup"><span data-stu-id="86de7-148">Requirement</span></span>| <span data-ttu-id="86de7-149">值</span><span class="sxs-lookup"><span data-stu-id="86de7-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="86de7-150">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="86de7-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86de7-151">1.0</span><span class="sxs-lookup"><span data-stu-id="86de7-151">1.0</span></span>|
|[<span data-ttu-id="86de7-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="86de7-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="86de7-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="86de7-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="86de7-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="86de7-154">SourceProperty :String</span></span>

<span data-ttu-id="86de7-155">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="86de7-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="86de7-156">类型：</span><span class="sxs-lookup"><span data-stu-id="86de7-156">Type:</span></span>

*   <span data-ttu-id="86de7-157">字符串</span><span class="sxs-lookup"><span data-stu-id="86de7-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="86de7-158">属性：</span><span class="sxs-lookup"><span data-stu-id="86de7-158">Properties:</span></span>

|<span data-ttu-id="86de7-159">名称</span><span class="sxs-lookup"><span data-stu-id="86de7-159">Name</span></span>| <span data-ttu-id="86de7-160">类型</span><span class="sxs-lookup"><span data-stu-id="86de7-160">Type</span></span>| <span data-ttu-id="86de7-161">描述</span><span class="sxs-lookup"><span data-stu-id="86de7-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="86de7-162">字符串</span><span class="sxs-lookup"><span data-stu-id="86de7-162">String</span></span>|<span data-ttu-id="86de7-163">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="86de7-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="86de7-164">String</span><span class="sxs-lookup"><span data-stu-id="86de7-164">String</span></span>|<span data-ttu-id="86de7-165">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="86de7-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="86de7-166">要求</span><span class="sxs-lookup"><span data-stu-id="86de7-166">Requirements</span></span>

|<span data-ttu-id="86de7-167">要求</span><span class="sxs-lookup"><span data-stu-id="86de7-167">Requirement</span></span>| <span data-ttu-id="86de7-168">值</span><span class="sxs-lookup"><span data-stu-id="86de7-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="86de7-169">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="86de7-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="86de7-170">1.0</span><span class="sxs-lookup"><span data-stu-id="86de7-170">1.0</span></span>|
|[<span data-ttu-id="86de7-171">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="86de7-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="86de7-172">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="86de7-172">Compose or read</span></span>|