---
title: Office 命名空间-要求集1。2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dc98d4c2da6e8f9ca294a6c686cf081478e1bb24
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450308"
---
# <a name="office"></a><span data-ttu-id="7ad1c-102">Office</span><span class="sxs-lookup"><span data-stu-id="7ad1c-102">Office</span></span>

<span data-ttu-id="7ad1c-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ad1c-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="7ad1c-105">Requirements</span></span>

|<span data-ttu-id="7ad1c-106">要求</span><span class="sxs-lookup"><span data-stu-id="7ad1c-106">Requirement</span></span>| <span data-ttu-id="7ad1c-107">值</span><span class="sxs-lookup"><span data-stu-id="7ad1c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ad1c-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7ad1c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ad1c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7ad1c-109">1.0</span></span>|
|[<span data-ttu-id="7ad1c-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7ad1c-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ad1c-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7ad1c-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="7ad1c-112">命名空间</span><span class="sxs-lookup"><span data-stu-id="7ad1c-112">Namespaces</span></span>

<span data-ttu-id="7ad1c-113">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="7ad1c-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="7ad1c-115">成员</span><span class="sxs-lookup"><span data-stu-id="7ad1c-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="7ad1c-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="7ad1c-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="7ad1c-117">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7ad1c-118">类型</span><span class="sxs-lookup"><span data-stu-id="7ad1c-118">Type</span></span>

*   <span data-ttu-id="7ad1c-119">String</span><span class="sxs-lookup"><span data-stu-id="7ad1c-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7ad1c-120">属性：</span><span class="sxs-lookup"><span data-stu-id="7ad1c-120">Properties:</span></span>

|<span data-ttu-id="7ad1c-121">名称</span><span class="sxs-lookup"><span data-stu-id="7ad1c-121">Name</span></span>| <span data-ttu-id="7ad1c-122">类型</span><span class="sxs-lookup"><span data-stu-id="7ad1c-122">Type</span></span>| <span data-ttu-id="7ad1c-123">描述</span><span class="sxs-lookup"><span data-stu-id="7ad1c-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7ad1c-124">字符串</span><span class="sxs-lookup"><span data-stu-id="7ad1c-124">String</span></span>|<span data-ttu-id="7ad1c-125">调用成功。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7ad1c-126">字符串</span><span class="sxs-lookup"><span data-stu-id="7ad1c-126">String</span></span>|<span data-ttu-id="7ad1c-127">调用失败。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7ad1c-128">Requirements</span><span class="sxs-lookup"><span data-stu-id="7ad1c-128">Requirements</span></span>

|<span data-ttu-id="7ad1c-129">要求</span><span class="sxs-lookup"><span data-stu-id="7ad1c-129">Requirement</span></span>| <span data-ttu-id="7ad1c-130">值</span><span class="sxs-lookup"><span data-stu-id="7ad1c-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ad1c-131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7ad1c-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ad1c-132">1.0</span><span class="sxs-lookup"><span data-stu-id="7ad1c-132">1.0</span></span>|
|[<span data-ttu-id="7ad1c-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7ad1c-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ad1c-134">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7ad1c-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="7ad1c-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="7ad1c-135">CoercionType :String</span></span>

<span data-ttu-id="7ad1c-136">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7ad1c-137">类型</span><span class="sxs-lookup"><span data-stu-id="7ad1c-137">Type</span></span>

*   <span data-ttu-id="7ad1c-138">String</span><span class="sxs-lookup"><span data-stu-id="7ad1c-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7ad1c-139">属性：</span><span class="sxs-lookup"><span data-stu-id="7ad1c-139">Properties:</span></span>

|<span data-ttu-id="7ad1c-140">名称</span><span class="sxs-lookup"><span data-stu-id="7ad1c-140">Name</span></span>| <span data-ttu-id="7ad1c-141">类型</span><span class="sxs-lookup"><span data-stu-id="7ad1c-141">Type</span></span>| <span data-ttu-id="7ad1c-142">描述</span><span class="sxs-lookup"><span data-stu-id="7ad1c-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7ad1c-143">字符串</span><span class="sxs-lookup"><span data-stu-id="7ad1c-143">String</span></span>|<span data-ttu-id="7ad1c-144">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7ad1c-145">字符串</span><span class="sxs-lookup"><span data-stu-id="7ad1c-145">String</span></span>|<span data-ttu-id="7ad1c-146">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7ad1c-147">Requirements</span><span class="sxs-lookup"><span data-stu-id="7ad1c-147">Requirements</span></span>

|<span data-ttu-id="7ad1c-148">要求</span><span class="sxs-lookup"><span data-stu-id="7ad1c-148">Requirement</span></span>| <span data-ttu-id="7ad1c-149">值</span><span class="sxs-lookup"><span data-stu-id="7ad1c-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ad1c-150">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7ad1c-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ad1c-151">1.0</span><span class="sxs-lookup"><span data-stu-id="7ad1c-151">1.0</span></span>|
|[<span data-ttu-id="7ad1c-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7ad1c-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ad1c-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7ad1c-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="7ad1c-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="7ad1c-154">SourceProperty :String</span></span>

<span data-ttu-id="7ad1c-155">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7ad1c-156">类型</span><span class="sxs-lookup"><span data-stu-id="7ad1c-156">Type</span></span>

*   <span data-ttu-id="7ad1c-157">String</span><span class="sxs-lookup"><span data-stu-id="7ad1c-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7ad1c-158">属性：</span><span class="sxs-lookup"><span data-stu-id="7ad1c-158">Properties:</span></span>

|<span data-ttu-id="7ad1c-159">名称</span><span class="sxs-lookup"><span data-stu-id="7ad1c-159">Name</span></span>| <span data-ttu-id="7ad1c-160">类型</span><span class="sxs-lookup"><span data-stu-id="7ad1c-160">Type</span></span>| <span data-ttu-id="7ad1c-161">描述</span><span class="sxs-lookup"><span data-stu-id="7ad1c-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7ad1c-162">字符串</span><span class="sxs-lookup"><span data-stu-id="7ad1c-162">String</span></span>|<span data-ttu-id="7ad1c-163">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7ad1c-164">String</span><span class="sxs-lookup"><span data-stu-id="7ad1c-164">String</span></span>|<span data-ttu-id="7ad1c-165">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="7ad1c-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7ad1c-166">Requirements</span><span class="sxs-lookup"><span data-stu-id="7ad1c-166">Requirements</span></span>

|<span data-ttu-id="7ad1c-167">要求</span><span class="sxs-lookup"><span data-stu-id="7ad1c-167">Requirement</span></span>| <span data-ttu-id="7ad1c-168">值</span><span class="sxs-lookup"><span data-stu-id="7ad1c-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ad1c-169">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="7ad1c-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7ad1c-170">1.0</span><span class="sxs-lookup"><span data-stu-id="7ad1c-170">1.0</span></span>|
|[<span data-ttu-id="7ad1c-171">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="7ad1c-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7ad1c-172">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="7ad1c-172">Compose or Read</span></span>|
