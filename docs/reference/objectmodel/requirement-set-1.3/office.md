---
title: Office 命名空间 - 要求集 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: ad08895719d809436216d2f0bb455260dbca3b1e
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067886"
---
# <a name="office"></a><span data-ttu-id="55734-102">Office</span><span class="sxs-lookup"><span data-stu-id="55734-102">Office</span></span>

<span data-ttu-id="55734-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="55734-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="55734-105">要求</span><span class="sxs-lookup"><span data-stu-id="55734-105">Requirements</span></span>

|<span data-ttu-id="55734-106">要求</span><span class="sxs-lookup"><span data-stu-id="55734-106">Requirement</span></span>| <span data-ttu-id="55734-107">值</span><span class="sxs-lookup"><span data-stu-id="55734-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="55734-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="55734-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55734-109">1.0</span><span class="sxs-lookup"><span data-stu-id="55734-109">1.0</span></span>|
|[<span data-ttu-id="55734-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="55734-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55734-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="55734-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="55734-112">命名空间</span><span class="sxs-lookup"><span data-stu-id="55734-112">Namespaces</span></span>

<span data-ttu-id="55734-113">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="55734-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="55734-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="55734-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="55734-115">成员</span><span class="sxs-lookup"><span data-stu-id="55734-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="55734-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="55734-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="55734-117">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="55734-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="55734-118">Type</span><span class="sxs-lookup"><span data-stu-id="55734-118">Type</span></span>

*   <span data-ttu-id="55734-119">String</span><span class="sxs-lookup"><span data-stu-id="55734-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55734-120">属性：</span><span class="sxs-lookup"><span data-stu-id="55734-120">Properties:</span></span>

|<span data-ttu-id="55734-121">名称</span><span class="sxs-lookup"><span data-stu-id="55734-121">Name</span></span>| <span data-ttu-id="55734-122">类型</span><span class="sxs-lookup"><span data-stu-id="55734-122">Type</span></span>| <span data-ttu-id="55734-123">说明</span><span class="sxs-lookup"><span data-stu-id="55734-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="55734-124">String</span><span class="sxs-lookup"><span data-stu-id="55734-124">String</span></span>|<span data-ttu-id="55734-125">调用成功。</span><span class="sxs-lookup"><span data-stu-id="55734-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="55734-126">String</span><span class="sxs-lookup"><span data-stu-id="55734-126">String</span></span>|<span data-ttu-id="55734-127">调用失败。</span><span class="sxs-lookup"><span data-stu-id="55734-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55734-128">Requirements</span><span class="sxs-lookup"><span data-stu-id="55734-128">Requirements</span></span>

|<span data-ttu-id="55734-129">要求</span><span class="sxs-lookup"><span data-stu-id="55734-129">Requirement</span></span>| <span data-ttu-id="55734-130">值</span><span class="sxs-lookup"><span data-stu-id="55734-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="55734-131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="55734-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55734-132">1.0</span><span class="sxs-lookup"><span data-stu-id="55734-132">1.0</span></span>|
|[<span data-ttu-id="55734-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="55734-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55734-134">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="55734-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="55734-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="55734-135">CoercionType :String</span></span>

<span data-ttu-id="55734-136">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="55734-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="55734-137">Type</span><span class="sxs-lookup"><span data-stu-id="55734-137">Type</span></span>

*   <span data-ttu-id="55734-138">String</span><span class="sxs-lookup"><span data-stu-id="55734-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55734-139">属性：</span><span class="sxs-lookup"><span data-stu-id="55734-139">Properties:</span></span>

|<span data-ttu-id="55734-140">名称</span><span class="sxs-lookup"><span data-stu-id="55734-140">Name</span></span>| <span data-ttu-id="55734-141">类型</span><span class="sxs-lookup"><span data-stu-id="55734-141">Type</span></span>| <span data-ttu-id="55734-142">说明</span><span class="sxs-lookup"><span data-stu-id="55734-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="55734-143">String</span><span class="sxs-lookup"><span data-stu-id="55734-143">String</span></span>|<span data-ttu-id="55734-144">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="55734-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="55734-145">字符串</span><span class="sxs-lookup"><span data-stu-id="55734-145">String</span></span>|<span data-ttu-id="55734-146">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="55734-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55734-147">要求</span><span class="sxs-lookup"><span data-stu-id="55734-147">Requirements</span></span>

|<span data-ttu-id="55734-148">要求</span><span class="sxs-lookup"><span data-stu-id="55734-148">Requirement</span></span>| <span data-ttu-id="55734-149">值</span><span class="sxs-lookup"><span data-stu-id="55734-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="55734-150">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="55734-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55734-151">1.0</span><span class="sxs-lookup"><span data-stu-id="55734-151">1.0</span></span>|
|[<span data-ttu-id="55734-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="55734-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55734-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="55734-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="55734-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="55734-154">SourceProperty :String</span></span>

<span data-ttu-id="55734-155">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="55734-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="55734-156">Type</span><span class="sxs-lookup"><span data-stu-id="55734-156">Type</span></span>

*   <span data-ttu-id="55734-157">String</span><span class="sxs-lookup"><span data-stu-id="55734-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55734-158">属性：</span><span class="sxs-lookup"><span data-stu-id="55734-158">Properties:</span></span>

|<span data-ttu-id="55734-159">名称</span><span class="sxs-lookup"><span data-stu-id="55734-159">Name</span></span>| <span data-ttu-id="55734-160">类型</span><span class="sxs-lookup"><span data-stu-id="55734-160">Type</span></span>| <span data-ttu-id="55734-161">说明</span><span class="sxs-lookup"><span data-stu-id="55734-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="55734-162">String</span><span class="sxs-lookup"><span data-stu-id="55734-162">String</span></span>|<span data-ttu-id="55734-163">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="55734-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="55734-164">String</span><span class="sxs-lookup"><span data-stu-id="55734-164">String</span></span>|<span data-ttu-id="55734-165">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="55734-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55734-166">要求</span><span class="sxs-lookup"><span data-stu-id="55734-166">Requirements</span></span>

|<span data-ttu-id="55734-167">要求</span><span class="sxs-lookup"><span data-stu-id="55734-167">Requirement</span></span>| <span data-ttu-id="55734-168">值</span><span class="sxs-lookup"><span data-stu-id="55734-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="55734-169">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="55734-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55734-170">1.0</span><span class="sxs-lookup"><span data-stu-id="55734-170">1.0</span></span>|
|[<span data-ttu-id="55734-171">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="55734-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="55734-172">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="55734-172">Compose or Read</span></span>|
