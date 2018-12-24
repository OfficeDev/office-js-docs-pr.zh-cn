---
title: Office 命令空间 - 要求集 1.4
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 7a86c550bd1f40c3db306c518165bc60b8bf0280
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433780"
---
# <a name="office"></a><span data-ttu-id="ba84f-102">Office</span><span class="sxs-lookup"><span data-stu-id="ba84f-102">Office</span></span>

<span data-ttu-id="ba84f-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="ba84f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ba84f-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="ba84f-105">Requirements</span></span>

|<span data-ttu-id="ba84f-106">要求</span><span class="sxs-lookup"><span data-stu-id="ba84f-106">Requirement</span></span>| <span data-ttu-id="ba84f-107">值</span><span class="sxs-lookup"><span data-stu-id="ba84f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba84f-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ba84f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba84f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ba84f-109">1.0</span></span>|
|[<span data-ttu-id="ba84f-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ba84f-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba84f-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ba84f-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="ba84f-112">命名空间</span><span class="sxs-lookup"><span data-stu-id="ba84f-112">Namespaces</span></span>

<span data-ttu-id="ba84f-113">[context](Office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="ba84f-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ba84f-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="ba84f-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ba84f-115">成员</span><span class="sxs-lookup"><span data-stu-id="ba84f-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ba84f-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ba84f-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="ba84f-117">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="ba84f-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ba84f-118">类型：</span><span class="sxs-lookup"><span data-stu-id="ba84f-118">Type:</span></span>

*   <span data-ttu-id="ba84f-119">字符串</span><span class="sxs-lookup"><span data-stu-id="ba84f-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba84f-120">属性：</span><span class="sxs-lookup"><span data-stu-id="ba84f-120">Properties:</span></span>

|<span data-ttu-id="ba84f-121">名称</span><span class="sxs-lookup"><span data-stu-id="ba84f-121">Name</span></span>| <span data-ttu-id="ba84f-122">类型</span><span class="sxs-lookup"><span data-stu-id="ba84f-122">Type</span></span>| <span data-ttu-id="ba84f-123">描述</span><span class="sxs-lookup"><span data-stu-id="ba84f-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ba84f-124">String</span><span class="sxs-lookup"><span data-stu-id="ba84f-124">String</span></span>|<span data-ttu-id="ba84f-125">调用成功。</span><span class="sxs-lookup"><span data-stu-id="ba84f-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ba84f-126">字符串</span><span class="sxs-lookup"><span data-stu-id="ba84f-126">String</span></span>|<span data-ttu-id="ba84f-127">调用失败。</span><span class="sxs-lookup"><span data-stu-id="ba84f-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba84f-128">要求</span><span class="sxs-lookup"><span data-stu-id="ba84f-128">Requirements</span></span>

|<span data-ttu-id="ba84f-129">要求</span><span class="sxs-lookup"><span data-stu-id="ba84f-129">Requirement</span></span>| <span data-ttu-id="ba84f-130">值</span><span class="sxs-lookup"><span data-stu-id="ba84f-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba84f-131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ba84f-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba84f-132">1.0</span><span class="sxs-lookup"><span data-stu-id="ba84f-132">1.0</span></span>|
|[<span data-ttu-id="ba84f-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ba84f-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba84f-134">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ba84f-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="ba84f-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ba84f-135">CoercionType :String</span></span>

<span data-ttu-id="ba84f-136">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="ba84f-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ba84f-137">类型：</span><span class="sxs-lookup"><span data-stu-id="ba84f-137">Type:</span></span>

*   <span data-ttu-id="ba84f-138">字符串</span><span class="sxs-lookup"><span data-stu-id="ba84f-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba84f-139">属性：</span><span class="sxs-lookup"><span data-stu-id="ba84f-139">Properties:</span></span>

|<span data-ttu-id="ba84f-140">名称</span><span class="sxs-lookup"><span data-stu-id="ba84f-140">Name</span></span>| <span data-ttu-id="ba84f-141">类型</span><span class="sxs-lookup"><span data-stu-id="ba84f-141">Type</span></span>| <span data-ttu-id="ba84f-142">描述</span><span class="sxs-lookup"><span data-stu-id="ba84f-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ba84f-143">String</span><span class="sxs-lookup"><span data-stu-id="ba84f-143">String</span></span>|<span data-ttu-id="ba84f-144">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="ba84f-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ba84f-145">字符串</span><span class="sxs-lookup"><span data-stu-id="ba84f-145">String</span></span>|<span data-ttu-id="ba84f-146">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="ba84f-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba84f-147">要求</span><span class="sxs-lookup"><span data-stu-id="ba84f-147">Requirements</span></span>

|<span data-ttu-id="ba84f-148">要求</span><span class="sxs-lookup"><span data-stu-id="ba84f-148">Requirement</span></span>| <span data-ttu-id="ba84f-149">值</span><span class="sxs-lookup"><span data-stu-id="ba84f-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba84f-150">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ba84f-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba84f-151">1.0</span><span class="sxs-lookup"><span data-stu-id="ba84f-151">1.0</span></span>|
|[<span data-ttu-id="ba84f-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ba84f-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba84f-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ba84f-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="ba84f-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ba84f-154">SourceProperty :String</span></span>

<span data-ttu-id="ba84f-155">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="ba84f-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ba84f-156">类型：</span><span class="sxs-lookup"><span data-stu-id="ba84f-156">Type:</span></span>

*   <span data-ttu-id="ba84f-157">字符串</span><span class="sxs-lookup"><span data-stu-id="ba84f-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba84f-158">属性：</span><span class="sxs-lookup"><span data-stu-id="ba84f-158">Properties:</span></span>

|<span data-ttu-id="ba84f-159">名称</span><span class="sxs-lookup"><span data-stu-id="ba84f-159">Name</span></span>| <span data-ttu-id="ba84f-160">类型</span><span class="sxs-lookup"><span data-stu-id="ba84f-160">Type</span></span>| <span data-ttu-id="ba84f-161">描述</span><span class="sxs-lookup"><span data-stu-id="ba84f-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ba84f-162">字符串</span><span class="sxs-lookup"><span data-stu-id="ba84f-162">String</span></span>|<span data-ttu-id="ba84f-163">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="ba84f-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ba84f-164">String</span><span class="sxs-lookup"><span data-stu-id="ba84f-164">String</span></span>|<span data-ttu-id="ba84f-165">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="ba84f-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba84f-166">要求</span><span class="sxs-lookup"><span data-stu-id="ba84f-166">Requirements</span></span>

|<span data-ttu-id="ba84f-167">要求</span><span class="sxs-lookup"><span data-stu-id="ba84f-167">Requirement</span></span>| <span data-ttu-id="ba84f-168">值</span><span class="sxs-lookup"><span data-stu-id="ba84f-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba84f-169">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ba84f-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba84f-170">1.0</span><span class="sxs-lookup"><span data-stu-id="ba84f-170">1.0</span></span>|
|[<span data-ttu-id="ba84f-171">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ba84f-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba84f-172">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ba84f-172">Compose or read</span></span>|