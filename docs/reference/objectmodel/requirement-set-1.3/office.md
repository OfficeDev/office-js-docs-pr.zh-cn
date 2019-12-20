---
title: Office 命名空间-要求集1。3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 3c6ddc34001f4d1622bc76d9bca1fbde9425be8b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814897"
---
# <a name="office"></a><span data-ttu-id="26541-102">Office</span><span class="sxs-lookup"><span data-stu-id="26541-102">Office</span></span>

<span data-ttu-id="26541-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="26541-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="26541-105">要求</span><span class="sxs-lookup"><span data-stu-id="26541-105">Requirements</span></span>

|<span data-ttu-id="26541-106">要求</span><span class="sxs-lookup"><span data-stu-id="26541-106">Requirement</span></span>| <span data-ttu-id="26541-107">值</span><span class="sxs-lookup"><span data-stu-id="26541-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="26541-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="26541-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="26541-109">1.1</span><span class="sxs-lookup"><span data-stu-id="26541-109">1.1</span></span>|
|[<span data-ttu-id="26541-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="26541-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26541-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="26541-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="26541-112">属性</span><span class="sxs-lookup"><span data-stu-id="26541-112">Properties</span></span>

| <span data-ttu-id="26541-113">属性</span><span class="sxs-lookup"><span data-stu-id="26541-113">Property</span></span> | <span data-ttu-id="26541-114">型号</span><span class="sxs-lookup"><span data-stu-id="26541-114">Modes</span></span> | <span data-ttu-id="26541-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="26541-115">Return type</span></span> | <span data-ttu-id="26541-116">最低</span><span class="sxs-lookup"><span data-stu-id="26541-116">Minimum</span></span><br><span data-ttu-id="26541-117">要求集</span><span class="sxs-lookup"><span data-stu-id="26541-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="26541-118">context</span><span class="sxs-lookup"><span data-stu-id="26541-118">context</span></span>](office.context.md) | <span data-ttu-id="26541-119">撰写</span><span class="sxs-lookup"><span data-stu-id="26541-119">Compose</span></span><br><span data-ttu-id="26541-120">读取</span><span class="sxs-lookup"><span data-stu-id="26541-120">Read</span></span> | [<span data-ttu-id="26541-121">Context</span><span class="sxs-lookup"><span data-stu-id="26541-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="26541-122">1.1</span><span class="sxs-lookup"><span data-stu-id="26541-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="26541-123">枚举</span><span class="sxs-lookup"><span data-stu-id="26541-123">Enumerations</span></span>

| <span data-ttu-id="26541-124">枚举</span><span class="sxs-lookup"><span data-stu-id="26541-124">Enumeration</span></span> | <span data-ttu-id="26541-125">型号</span><span class="sxs-lookup"><span data-stu-id="26541-125">Modes</span></span> | <span data-ttu-id="26541-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="26541-126">Return type</span></span> | <span data-ttu-id="26541-127">最低</span><span class="sxs-lookup"><span data-stu-id="26541-127">Minimum</span></span><br><span data-ttu-id="26541-128">要求集</span><span class="sxs-lookup"><span data-stu-id="26541-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="26541-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="26541-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="26541-130">撰写</span><span class="sxs-lookup"><span data-stu-id="26541-130">Compose</span></span><br><span data-ttu-id="26541-131">读取</span><span class="sxs-lookup"><span data-stu-id="26541-131">Read</span></span> | <span data-ttu-id="26541-132">String</span><span class="sxs-lookup"><span data-stu-id="26541-132">String</span></span> | [<span data-ttu-id="26541-133">1.1</span><span class="sxs-lookup"><span data-stu-id="26541-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="26541-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="26541-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="26541-135">撰写</span><span class="sxs-lookup"><span data-stu-id="26541-135">Compose</span></span><br><span data-ttu-id="26541-136">读取</span><span class="sxs-lookup"><span data-stu-id="26541-136">Read</span></span> | <span data-ttu-id="26541-137">String</span><span class="sxs-lookup"><span data-stu-id="26541-137">String</span></span> | [<span data-ttu-id="26541-138">1.1</span><span class="sxs-lookup"><span data-stu-id="26541-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="26541-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="26541-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="26541-140">撰写</span><span class="sxs-lookup"><span data-stu-id="26541-140">Compose</span></span><br><span data-ttu-id="26541-141">读取</span><span class="sxs-lookup"><span data-stu-id="26541-141">Read</span></span> | <span data-ttu-id="26541-142">String</span><span class="sxs-lookup"><span data-stu-id="26541-142">String</span></span> | [<span data-ttu-id="26541-143">1.1</span><span class="sxs-lookup"><span data-stu-id="26541-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="26541-144">命名空间</span><span class="sxs-lookup"><span data-stu-id="26541-144">Namespaces</span></span>

<span data-ttu-id="26541-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="26541-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="26541-146">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="26541-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="26541-147">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="26541-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="26541-148">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="26541-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="26541-149">类型</span><span class="sxs-lookup"><span data-stu-id="26541-149">Type</span></span>

*   <span data-ttu-id="26541-150">String</span><span class="sxs-lookup"><span data-stu-id="26541-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="26541-151">属性：</span><span class="sxs-lookup"><span data-stu-id="26541-151">Properties:</span></span>

|<span data-ttu-id="26541-152">名称</span><span class="sxs-lookup"><span data-stu-id="26541-152">Name</span></span>| <span data-ttu-id="26541-153">类型</span><span class="sxs-lookup"><span data-stu-id="26541-153">Type</span></span>| <span data-ttu-id="26541-154">说明</span><span class="sxs-lookup"><span data-stu-id="26541-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="26541-155">String</span><span class="sxs-lookup"><span data-stu-id="26541-155">String</span></span>|<span data-ttu-id="26541-156">调用成功。</span><span class="sxs-lookup"><span data-stu-id="26541-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="26541-157">String</span><span class="sxs-lookup"><span data-stu-id="26541-157">String</span></span>|<span data-ttu-id="26541-158">调用失败。</span><span class="sxs-lookup"><span data-stu-id="26541-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26541-159">要求</span><span class="sxs-lookup"><span data-stu-id="26541-159">Requirements</span></span>

|<span data-ttu-id="26541-160">要求</span><span class="sxs-lookup"><span data-stu-id="26541-160">Requirement</span></span>| <span data-ttu-id="26541-161">值</span><span class="sxs-lookup"><span data-stu-id="26541-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="26541-162">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="26541-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="26541-163">1.1</span><span class="sxs-lookup"><span data-stu-id="26541-163">1.1</span></span>|
|[<span data-ttu-id="26541-164">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="26541-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26541-165">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="26541-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="26541-166">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="26541-166">CoercionType: String</span></span>

<span data-ttu-id="26541-167">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="26541-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="26541-168">类型</span><span class="sxs-lookup"><span data-stu-id="26541-168">Type</span></span>

*   <span data-ttu-id="26541-169">String</span><span class="sxs-lookup"><span data-stu-id="26541-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="26541-170">属性：</span><span class="sxs-lookup"><span data-stu-id="26541-170">Properties:</span></span>

|<span data-ttu-id="26541-171">名称</span><span class="sxs-lookup"><span data-stu-id="26541-171">Name</span></span>| <span data-ttu-id="26541-172">类型</span><span class="sxs-lookup"><span data-stu-id="26541-172">Type</span></span>| <span data-ttu-id="26541-173">说明</span><span class="sxs-lookup"><span data-stu-id="26541-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="26541-174">String</span><span class="sxs-lookup"><span data-stu-id="26541-174">String</span></span>|<span data-ttu-id="26541-175">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="26541-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="26541-176">String</span><span class="sxs-lookup"><span data-stu-id="26541-176">String</span></span>|<span data-ttu-id="26541-177">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="26541-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26541-178">要求</span><span class="sxs-lookup"><span data-stu-id="26541-178">Requirements</span></span>

|<span data-ttu-id="26541-179">要求</span><span class="sxs-lookup"><span data-stu-id="26541-179">Requirement</span></span>| <span data-ttu-id="26541-180">值</span><span class="sxs-lookup"><span data-stu-id="26541-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="26541-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="26541-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="26541-182">1.1</span><span class="sxs-lookup"><span data-stu-id="26541-182">1.1</span></span>|
|[<span data-ttu-id="26541-183">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="26541-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26541-184">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="26541-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="26541-185">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="26541-185">SourceProperty: String</span></span>

<span data-ttu-id="26541-186">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="26541-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="26541-187">类型</span><span class="sxs-lookup"><span data-stu-id="26541-187">Type</span></span>

*   <span data-ttu-id="26541-188">String</span><span class="sxs-lookup"><span data-stu-id="26541-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="26541-189">属性：</span><span class="sxs-lookup"><span data-stu-id="26541-189">Properties:</span></span>

|<span data-ttu-id="26541-190">名称</span><span class="sxs-lookup"><span data-stu-id="26541-190">Name</span></span>| <span data-ttu-id="26541-191">类型</span><span class="sxs-lookup"><span data-stu-id="26541-191">Type</span></span>| <span data-ttu-id="26541-192">说明</span><span class="sxs-lookup"><span data-stu-id="26541-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="26541-193">String</span><span class="sxs-lookup"><span data-stu-id="26541-193">String</span></span>|<span data-ttu-id="26541-194">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="26541-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="26541-195">String</span><span class="sxs-lookup"><span data-stu-id="26541-195">String</span></span>|<span data-ttu-id="26541-196">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="26541-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26541-197">要求</span><span class="sxs-lookup"><span data-stu-id="26541-197">Requirements</span></span>

|<span data-ttu-id="26541-198">要求</span><span class="sxs-lookup"><span data-stu-id="26541-198">Requirement</span></span>| <span data-ttu-id="26541-199">值</span><span class="sxs-lookup"><span data-stu-id="26541-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="26541-200">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="26541-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="26541-201">1.1</span><span class="sxs-lookup"><span data-stu-id="26541-201">1.1</span></span>|
|[<span data-ttu-id="26541-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="26541-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26541-203">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="26541-203">Compose or Read</span></span>|
