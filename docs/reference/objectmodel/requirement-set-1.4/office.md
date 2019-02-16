---
title: Office 命令空间 - 要求集 1.4
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: af5e05e2243c0132018bc4eba7006f9a5aad4099
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067949"
---
# <a name="office"></a><span data-ttu-id="c0184-102">Office</span><span class="sxs-lookup"><span data-stu-id="c0184-102">Office</span></span>

<span data-ttu-id="c0184-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="c0184-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0184-105">要求</span><span class="sxs-lookup"><span data-stu-id="c0184-105">Requirements</span></span>

|<span data-ttu-id="c0184-106">要求</span><span class="sxs-lookup"><span data-stu-id="c0184-106">Requirement</span></span>| <span data-ttu-id="c0184-107">值</span><span class="sxs-lookup"><span data-stu-id="c0184-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0184-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0184-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0184-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c0184-109">1.0</span></span>|
|[<span data-ttu-id="c0184-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0184-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c0184-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0184-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="c0184-112">命名空间</span><span class="sxs-lookup"><span data-stu-id="c0184-112">Namespaces</span></span>

<span data-ttu-id="c0184-113">[context](Office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="c0184-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c0184-114">[MailboxEnums](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="c0184-114">[MailboxEnums](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c0184-115">成员</span><span class="sxs-lookup"><span data-stu-id="c0184-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c0184-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c0184-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="c0184-117">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="c0184-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c0184-118">Type</span><span class="sxs-lookup"><span data-stu-id="c0184-118">Type</span></span>

*   <span data-ttu-id="c0184-119">String</span><span class="sxs-lookup"><span data-stu-id="c0184-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c0184-120">属性：</span><span class="sxs-lookup"><span data-stu-id="c0184-120">Properties:</span></span>

|<span data-ttu-id="c0184-121">名称</span><span class="sxs-lookup"><span data-stu-id="c0184-121">Name</span></span>| <span data-ttu-id="c0184-122">类型</span><span class="sxs-lookup"><span data-stu-id="c0184-122">Type</span></span>| <span data-ttu-id="c0184-123">说明</span><span class="sxs-lookup"><span data-stu-id="c0184-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c0184-124">String</span><span class="sxs-lookup"><span data-stu-id="c0184-124">String</span></span>|<span data-ttu-id="c0184-125">调用成功。</span><span class="sxs-lookup"><span data-stu-id="c0184-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c0184-126">String</span><span class="sxs-lookup"><span data-stu-id="c0184-126">String</span></span>|<span data-ttu-id="c0184-127">调用失败。</span><span class="sxs-lookup"><span data-stu-id="c0184-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c0184-128">Requirements</span><span class="sxs-lookup"><span data-stu-id="c0184-128">Requirements</span></span>

|<span data-ttu-id="c0184-129">要求</span><span class="sxs-lookup"><span data-stu-id="c0184-129">Requirement</span></span>| <span data-ttu-id="c0184-130">值</span><span class="sxs-lookup"><span data-stu-id="c0184-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0184-131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0184-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0184-132">1.0</span><span class="sxs-lookup"><span data-stu-id="c0184-132">1.0</span></span>|
|[<span data-ttu-id="c0184-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0184-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c0184-134">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0184-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="c0184-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c0184-135">CoercionType :String</span></span>

<span data-ttu-id="c0184-136">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="c0184-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c0184-137">Type</span><span class="sxs-lookup"><span data-stu-id="c0184-137">Type</span></span>

*   <span data-ttu-id="c0184-138">String</span><span class="sxs-lookup"><span data-stu-id="c0184-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c0184-139">属性：</span><span class="sxs-lookup"><span data-stu-id="c0184-139">Properties:</span></span>

|<span data-ttu-id="c0184-140">名称</span><span class="sxs-lookup"><span data-stu-id="c0184-140">Name</span></span>| <span data-ttu-id="c0184-141">类型</span><span class="sxs-lookup"><span data-stu-id="c0184-141">Type</span></span>| <span data-ttu-id="c0184-142">说明</span><span class="sxs-lookup"><span data-stu-id="c0184-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c0184-143">String</span><span class="sxs-lookup"><span data-stu-id="c0184-143">String</span></span>|<span data-ttu-id="c0184-144">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="c0184-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c0184-145">字符串</span><span class="sxs-lookup"><span data-stu-id="c0184-145">String</span></span>|<span data-ttu-id="c0184-146">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="c0184-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c0184-147">要求</span><span class="sxs-lookup"><span data-stu-id="c0184-147">Requirements</span></span>

|<span data-ttu-id="c0184-148">要求</span><span class="sxs-lookup"><span data-stu-id="c0184-148">Requirement</span></span>| <span data-ttu-id="c0184-149">值</span><span class="sxs-lookup"><span data-stu-id="c0184-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0184-150">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0184-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0184-151">1.0</span><span class="sxs-lookup"><span data-stu-id="c0184-151">1.0</span></span>|
|[<span data-ttu-id="c0184-152">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0184-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c0184-153">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0184-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="c0184-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c0184-154">SourceProperty :String</span></span>

<span data-ttu-id="c0184-155">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="c0184-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c0184-156">Type</span><span class="sxs-lookup"><span data-stu-id="c0184-156">Type</span></span>

*   <span data-ttu-id="c0184-157">String</span><span class="sxs-lookup"><span data-stu-id="c0184-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c0184-158">属性：</span><span class="sxs-lookup"><span data-stu-id="c0184-158">Properties:</span></span>

|<span data-ttu-id="c0184-159">名称</span><span class="sxs-lookup"><span data-stu-id="c0184-159">Name</span></span>| <span data-ttu-id="c0184-160">类型</span><span class="sxs-lookup"><span data-stu-id="c0184-160">Type</span></span>| <span data-ttu-id="c0184-161">说明</span><span class="sxs-lookup"><span data-stu-id="c0184-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c0184-162">String</span><span class="sxs-lookup"><span data-stu-id="c0184-162">String</span></span>|<span data-ttu-id="c0184-163">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="c0184-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c0184-164">String</span><span class="sxs-lookup"><span data-stu-id="c0184-164">String</span></span>|<span data-ttu-id="c0184-165">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="c0184-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c0184-166">要求</span><span class="sxs-lookup"><span data-stu-id="c0184-166">Requirements</span></span>

|<span data-ttu-id="c0184-167">要求</span><span class="sxs-lookup"><span data-stu-id="c0184-167">Requirement</span></span>| <span data-ttu-id="c0184-168">值</span><span class="sxs-lookup"><span data-stu-id="c0184-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0184-169">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c0184-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0184-170">1.0</span><span class="sxs-lookup"><span data-stu-id="c0184-170">1.0</span></span>|
|[<span data-ttu-id="c0184-171">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c0184-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c0184-172">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c0184-172">Compose or Read</span></span>|
