---
title: Office 命名空间-要求集1。2
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 25322842d82c1db7228a67c268facb6945168222
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696293"
---
# <a name="office"></a><span data-ttu-id="436c7-102">Office</span><span class="sxs-lookup"><span data-stu-id="436c7-102">Office</span></span>

<span data-ttu-id="436c7-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="436c7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="436c7-105">要求</span><span class="sxs-lookup"><span data-stu-id="436c7-105">Requirements</span></span>

|<span data-ttu-id="436c7-106">要求</span><span class="sxs-lookup"><span data-stu-id="436c7-106">Requirement</span></span>| <span data-ttu-id="436c7-107">值</span><span class="sxs-lookup"><span data-stu-id="436c7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="436c7-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="436c7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="436c7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="436c7-109">1.0</span></span>|
|[<span data-ttu-id="436c7-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="436c7-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="436c7-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="436c7-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="436c7-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="436c7-112">Members and methods</span></span>

| <span data-ttu-id="436c7-113">成员</span><span class="sxs-lookup"><span data-stu-id="436c7-113">Member</span></span> | <span data-ttu-id="436c7-114">类型</span><span class="sxs-lookup"><span data-stu-id="436c7-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="436c7-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="436c7-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="436c7-116">Member</span><span class="sxs-lookup"><span data-stu-id="436c7-116">Member</span></span> |
| [<span data-ttu-id="436c7-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="436c7-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="436c7-118">Member</span><span class="sxs-lookup"><span data-stu-id="436c7-118">Member</span></span> |
| [<span data-ttu-id="436c7-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="436c7-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="436c7-120">成员</span><span class="sxs-lookup"><span data-stu-id="436c7-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="436c7-121">命名空间</span><span class="sxs-lookup"><span data-stu-id="436c7-121">Namespaces</span></span>

<span data-ttu-id="436c7-122">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="436c7-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="436c7-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): 包含多个`ItemType`枚举, 例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="436c7-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="436c7-124">Members</span><span class="sxs-lookup"><span data-stu-id="436c7-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="436c7-125">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="436c7-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="436c7-126">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="436c7-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="436c7-127">类型</span><span class="sxs-lookup"><span data-stu-id="436c7-127">Type</span></span>

*   <span data-ttu-id="436c7-128">String</span><span class="sxs-lookup"><span data-stu-id="436c7-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="436c7-129">属性：</span><span class="sxs-lookup"><span data-stu-id="436c7-129">Properties:</span></span>

|<span data-ttu-id="436c7-130">名称</span><span class="sxs-lookup"><span data-stu-id="436c7-130">Name</span></span>| <span data-ttu-id="436c7-131">类型</span><span class="sxs-lookup"><span data-stu-id="436c7-131">Type</span></span>| <span data-ttu-id="436c7-132">说明</span><span class="sxs-lookup"><span data-stu-id="436c7-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="436c7-133">String</span><span class="sxs-lookup"><span data-stu-id="436c7-133">String</span></span>|<span data-ttu-id="436c7-134">调用成功。</span><span class="sxs-lookup"><span data-stu-id="436c7-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="436c7-135">String</span><span class="sxs-lookup"><span data-stu-id="436c7-135">String</span></span>|<span data-ttu-id="436c7-136">调用失败。</span><span class="sxs-lookup"><span data-stu-id="436c7-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="436c7-137">要求</span><span class="sxs-lookup"><span data-stu-id="436c7-137">Requirements</span></span>

|<span data-ttu-id="436c7-138">要求</span><span class="sxs-lookup"><span data-stu-id="436c7-138">Requirement</span></span>| <span data-ttu-id="436c7-139">值</span><span class="sxs-lookup"><span data-stu-id="436c7-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="436c7-140">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="436c7-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="436c7-141">1.0</span><span class="sxs-lookup"><span data-stu-id="436c7-141">1.0</span></span>|
|[<span data-ttu-id="436c7-142">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="436c7-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="436c7-143">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="436c7-143">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="436c7-144">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="436c7-144">CoercionType: String</span></span>

<span data-ttu-id="436c7-145">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="436c7-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="436c7-146">类型</span><span class="sxs-lookup"><span data-stu-id="436c7-146">Type</span></span>

*   <span data-ttu-id="436c7-147">String</span><span class="sxs-lookup"><span data-stu-id="436c7-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="436c7-148">属性：</span><span class="sxs-lookup"><span data-stu-id="436c7-148">Properties:</span></span>

|<span data-ttu-id="436c7-149">名称</span><span class="sxs-lookup"><span data-stu-id="436c7-149">Name</span></span>| <span data-ttu-id="436c7-150">类型</span><span class="sxs-lookup"><span data-stu-id="436c7-150">Type</span></span>| <span data-ttu-id="436c7-151">说明</span><span class="sxs-lookup"><span data-stu-id="436c7-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="436c7-152">String</span><span class="sxs-lookup"><span data-stu-id="436c7-152">String</span></span>|<span data-ttu-id="436c7-153">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="436c7-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="436c7-154">String</span><span class="sxs-lookup"><span data-stu-id="436c7-154">String</span></span>|<span data-ttu-id="436c7-155">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="436c7-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="436c7-156">要求</span><span class="sxs-lookup"><span data-stu-id="436c7-156">Requirements</span></span>

|<span data-ttu-id="436c7-157">要求</span><span class="sxs-lookup"><span data-stu-id="436c7-157">Requirement</span></span>| <span data-ttu-id="436c7-158">值</span><span class="sxs-lookup"><span data-stu-id="436c7-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="436c7-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="436c7-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="436c7-160">1.0</span><span class="sxs-lookup"><span data-stu-id="436c7-160">1.0</span></span>|
|[<span data-ttu-id="436c7-161">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="436c7-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="436c7-162">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="436c7-162">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="436c7-163">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="436c7-163">SourceProperty: String</span></span>

<span data-ttu-id="436c7-164">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="436c7-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="436c7-165">类型</span><span class="sxs-lookup"><span data-stu-id="436c7-165">Type</span></span>

*   <span data-ttu-id="436c7-166">String</span><span class="sxs-lookup"><span data-stu-id="436c7-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="436c7-167">属性：</span><span class="sxs-lookup"><span data-stu-id="436c7-167">Properties:</span></span>

|<span data-ttu-id="436c7-168">名称</span><span class="sxs-lookup"><span data-stu-id="436c7-168">Name</span></span>| <span data-ttu-id="436c7-169">类型</span><span class="sxs-lookup"><span data-stu-id="436c7-169">Type</span></span>| <span data-ttu-id="436c7-170">说明</span><span class="sxs-lookup"><span data-stu-id="436c7-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="436c7-171">String</span><span class="sxs-lookup"><span data-stu-id="436c7-171">String</span></span>|<span data-ttu-id="436c7-172">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="436c7-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="436c7-173">String</span><span class="sxs-lookup"><span data-stu-id="436c7-173">String</span></span>|<span data-ttu-id="436c7-174">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="436c7-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="436c7-175">要求</span><span class="sxs-lookup"><span data-stu-id="436c7-175">Requirements</span></span>

|<span data-ttu-id="436c7-176">要求</span><span class="sxs-lookup"><span data-stu-id="436c7-176">Requirement</span></span>| <span data-ttu-id="436c7-177">值</span><span class="sxs-lookup"><span data-stu-id="436c7-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="436c7-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="436c7-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="436c7-179">1.0</span><span class="sxs-lookup"><span data-stu-id="436c7-179">1.0</span></span>|
|[<span data-ttu-id="436c7-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="436c7-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="436c7-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="436c7-181">Compose or Read</span></span>|
