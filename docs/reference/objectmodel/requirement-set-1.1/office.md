---
title: Office 命名空间 - 要求设置 1.1
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: af2a48d5bc943d4f443c32777fefaf8ed4a30032
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457430"
---
# <a name="office"></a><span data-ttu-id="57f9e-102">Office</span><span class="sxs-lookup"><span data-stu-id="57f9e-102">Office</span></span>

<span data-ttu-id="57f9e-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="57f9e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="57f9e-105">要求</span><span class="sxs-lookup"><span data-stu-id="57f9e-105">Requirements</span></span>

|<span data-ttu-id="57f9e-106">要求</span><span class="sxs-lookup"><span data-stu-id="57f9e-106">Requirement</span></span>| <span data-ttu-id="57f9e-107">值</span><span class="sxs-lookup"><span data-stu-id="57f9e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="57f9e-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="57f9e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57f9e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="57f9e-109">1.0</span></span>|
|[<span data-ttu-id="57f9e-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="57f9e-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="57f9e-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="57f9e-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="57f9e-112">命名空间</span><span class="sxs-lookup"><span data-stu-id="57f9e-112">Namespaces</span></span>

<span data-ttu-id="57f9e-113">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="57f9e-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="57f9e-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="57f9e-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="57f9e-115">成员</span><span class="sxs-lookup"><span data-stu-id="57f9e-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="57f9e-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="57f9e-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="57f9e-117">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="57f9e-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="57f9e-118">类型：</span><span class="sxs-lookup"><span data-stu-id="57f9e-118">Type:</span></span>

*   <span data-ttu-id="57f9e-119">字符串</span><span class="sxs-lookup"><span data-stu-id="57f9e-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="57f9e-120">属性：</span><span class="sxs-lookup"><span data-stu-id="57f9e-120">Properties:</span></span>

|<span data-ttu-id="57f9e-121">名称</span><span class="sxs-lookup"><span data-stu-id="57f9e-121">Name</span></span>| <span data-ttu-id="57f9e-122">类型</span><span class="sxs-lookup"><span data-stu-id="57f9e-122">Type</span></span>| <span data-ttu-id="57f9e-123">描述</span><span class="sxs-lookup"><span data-stu-id="57f9e-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="57f9e-124">String</span><span class="sxs-lookup"><span data-stu-id="57f9e-124">String</span></span>|<span data-ttu-id="57f9e-125">调用成功。</span><span class="sxs-lookup"><span data-stu-id="57f9e-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="57f9e-126">字符串</span><span class="sxs-lookup"><span data-stu-id="57f9e-126">String</span></span>|<span data-ttu-id="57f9e-127">调用失败。</span><span class="sxs-lookup"><span data-stu-id="57f9e-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57f9e-128">要求</span><span class="sxs-lookup"><span data-stu-id="57f9e-128">Requirements</span></span>

|<span data-ttu-id="57f9e-129">要求</span><span class="sxs-lookup"><span data-stu-id="57f9e-129">Requirement</span></span>| <span data-ttu-id="57f9e-130">值</span><span class="sxs-lookup"><span data-stu-id="57f9e-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="57f9e-131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="57f9e-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57f9e-132">1.0</span><span class="sxs-lookup"><span data-stu-id="57f9e-132">1.0</span></span>|
|[<span data-ttu-id="57f9e-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="57f9e-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="57f9e-134">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="57f9e-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="57f9e-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="57f9e-135">CoercionType :String</span></span>

<span data-ttu-id="57f9e-136">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="57f9e-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="57f9e-137">类型：</span><span class="sxs-lookup"><span data-stu-id="57f9e-137">Type:</span></span>

*   <span data-ttu-id="57f9e-138">字符串</span><span class="sxs-lookup"><span data-stu-id="57f9e-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="57f9e-139">属性：</span><span class="sxs-lookup"><span data-stu-id="57f9e-139">Properties:</span></span>

|<span data-ttu-id="57f9e-140">名称</span><span class="sxs-lookup"><span data-stu-id="57f9e-140">Name</span></span>| <span data-ttu-id="57f9e-141">类型</span><span class="sxs-lookup"><span data-stu-id="57f9e-141">Type</span></span>| <span data-ttu-id="57f9e-142">描述</span><span class="sxs-lookup"><span data-stu-id="57f9e-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="57f9e-143">String</span><span class="sxs-lookup"><span data-stu-id="57f9e-143">String</span></span>|<span data-ttu-id="57f9e-144">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="57f9e-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="57f9e-145">字符串</span><span class="sxs-lookup"><span data-stu-id="57f9e-145">String</span></span>|<span data-ttu-id="57f9e-146">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="57f9e-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57f9e-147">要求</span><span class="sxs-lookup"><span data-stu-id="57f9e-147">Requirements</span></span>

|<span data-ttu-id="57f9e-148">要求</span><span class="sxs-lookup"><span data-stu-id="57f9e-148">Requirement</span></span>| <span data-ttu-id="57f9e-149">值</span><span class="sxs-lookup"><span data-stu-id="57f9e-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="57f9e-150">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="57f9e-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57f9e-151">1.0</span><span class="sxs-lookup"><span data-stu-id="57f9e-151">1.0</span></span>|
|[<span data-ttu-id="57f9e-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="57f9e-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="57f9e-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="57f9e-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="57f9e-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="57f9e-154">SourceProperty :String</span></span>

<span data-ttu-id="57f9e-155">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="57f9e-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="57f9e-156">类型：</span><span class="sxs-lookup"><span data-stu-id="57f9e-156">Type:</span></span>

*   <span data-ttu-id="57f9e-157">字符串</span><span class="sxs-lookup"><span data-stu-id="57f9e-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="57f9e-158">属性：</span><span class="sxs-lookup"><span data-stu-id="57f9e-158">Properties:</span></span>

|<span data-ttu-id="57f9e-159">名称</span><span class="sxs-lookup"><span data-stu-id="57f9e-159">Name</span></span>| <span data-ttu-id="57f9e-160">类型</span><span class="sxs-lookup"><span data-stu-id="57f9e-160">Type</span></span>| <span data-ttu-id="57f9e-161">描述</span><span class="sxs-lookup"><span data-stu-id="57f9e-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="57f9e-162">字符串</span><span class="sxs-lookup"><span data-stu-id="57f9e-162">String</span></span>|<span data-ttu-id="57f9e-163">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="57f9e-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="57f9e-164">String</span><span class="sxs-lookup"><span data-stu-id="57f9e-164">String</span></span>|<span data-ttu-id="57f9e-165">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="57f9e-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57f9e-166">要求</span><span class="sxs-lookup"><span data-stu-id="57f9e-166">Requirements</span></span>

|<span data-ttu-id="57f9e-167">要求</span><span class="sxs-lookup"><span data-stu-id="57f9e-167">Requirement</span></span>| <span data-ttu-id="57f9e-168">值</span><span class="sxs-lookup"><span data-stu-id="57f9e-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="57f9e-169">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="57f9e-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57f9e-170">1.0</span><span class="sxs-lookup"><span data-stu-id="57f9e-170">1.0</span></span>|
|[<span data-ttu-id="57f9e-171">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="57f9e-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="57f9e-172">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="57f9e-172">Compose or read</span></span>|