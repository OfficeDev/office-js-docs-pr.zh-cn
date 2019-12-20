---
title: Office 命名空间-要求集1。2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 714bbd6dfdcd47687a2309c24fd666168aed6556
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814325"
---
# <a name="office"></a><span data-ttu-id="05300-102">Office</span><span class="sxs-lookup"><span data-stu-id="05300-102">Office</span></span>

<span data-ttu-id="05300-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="05300-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="05300-105">要求</span><span class="sxs-lookup"><span data-stu-id="05300-105">Requirements</span></span>

|<span data-ttu-id="05300-106">要求</span><span class="sxs-lookup"><span data-stu-id="05300-106">Requirement</span></span>| <span data-ttu-id="05300-107">值</span><span class="sxs-lookup"><span data-stu-id="05300-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="05300-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="05300-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05300-109">1.1</span><span class="sxs-lookup"><span data-stu-id="05300-109">1.1</span></span>|
|[<span data-ttu-id="05300-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="05300-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05300-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="05300-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="05300-112">属性</span><span class="sxs-lookup"><span data-stu-id="05300-112">Properties</span></span>

| <span data-ttu-id="05300-113">属性</span><span class="sxs-lookup"><span data-stu-id="05300-113">Property</span></span> | <span data-ttu-id="05300-114">型号</span><span class="sxs-lookup"><span data-stu-id="05300-114">Modes</span></span> | <span data-ttu-id="05300-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="05300-115">Return type</span></span> | <span data-ttu-id="05300-116">最低</span><span class="sxs-lookup"><span data-stu-id="05300-116">Minimum</span></span><br><span data-ttu-id="05300-117">要求集</span><span class="sxs-lookup"><span data-stu-id="05300-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="05300-118">context</span><span class="sxs-lookup"><span data-stu-id="05300-118">context</span></span>](office.context.md) | <span data-ttu-id="05300-119">撰写</span><span class="sxs-lookup"><span data-stu-id="05300-119">Compose</span></span><br><span data-ttu-id="05300-120">读取</span><span class="sxs-lookup"><span data-stu-id="05300-120">Read</span></span> | [<span data-ttu-id="05300-121">Context</span><span class="sxs-lookup"><span data-stu-id="05300-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2) | [<span data-ttu-id="05300-122">1.1</span><span class="sxs-lookup"><span data-stu-id="05300-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="05300-123">枚举</span><span class="sxs-lookup"><span data-stu-id="05300-123">Enumerations</span></span>

| <span data-ttu-id="05300-124">枚举</span><span class="sxs-lookup"><span data-stu-id="05300-124">Enumeration</span></span> | <span data-ttu-id="05300-125">型号</span><span class="sxs-lookup"><span data-stu-id="05300-125">Modes</span></span> | <span data-ttu-id="05300-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="05300-126">Return type</span></span> | <span data-ttu-id="05300-127">最低</span><span class="sxs-lookup"><span data-stu-id="05300-127">Minimum</span></span><br><span data-ttu-id="05300-128">要求集</span><span class="sxs-lookup"><span data-stu-id="05300-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="05300-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="05300-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="05300-130">撰写</span><span class="sxs-lookup"><span data-stu-id="05300-130">Compose</span></span><br><span data-ttu-id="05300-131">读取</span><span class="sxs-lookup"><span data-stu-id="05300-131">Read</span></span> | <span data-ttu-id="05300-132">String</span><span class="sxs-lookup"><span data-stu-id="05300-132">String</span></span> | [<span data-ttu-id="05300-133">1.1</span><span class="sxs-lookup"><span data-stu-id="05300-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05300-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="05300-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="05300-135">撰写</span><span class="sxs-lookup"><span data-stu-id="05300-135">Compose</span></span><br><span data-ttu-id="05300-136">读取</span><span class="sxs-lookup"><span data-stu-id="05300-136">Read</span></span> | <span data-ttu-id="05300-137">String</span><span class="sxs-lookup"><span data-stu-id="05300-137">String</span></span> | [<span data-ttu-id="05300-138">1.1</span><span class="sxs-lookup"><span data-stu-id="05300-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05300-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="05300-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="05300-140">撰写</span><span class="sxs-lookup"><span data-stu-id="05300-140">Compose</span></span><br><span data-ttu-id="05300-141">读取</span><span class="sxs-lookup"><span data-stu-id="05300-141">Read</span></span> | <span data-ttu-id="05300-142">String</span><span class="sxs-lookup"><span data-stu-id="05300-142">String</span></span> | [<span data-ttu-id="05300-143">1.1</span><span class="sxs-lookup"><span data-stu-id="05300-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="05300-144">命名空间</span><span class="sxs-lookup"><span data-stu-id="05300-144">Namespaces</span></span>

<span data-ttu-id="05300-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="05300-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="05300-146">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="05300-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="05300-147">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="05300-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="05300-148">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="05300-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="05300-149">类型</span><span class="sxs-lookup"><span data-stu-id="05300-149">Type</span></span>

*   <span data-ttu-id="05300-150">String</span><span class="sxs-lookup"><span data-stu-id="05300-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="05300-151">属性：</span><span class="sxs-lookup"><span data-stu-id="05300-151">Properties:</span></span>

|<span data-ttu-id="05300-152">名称</span><span class="sxs-lookup"><span data-stu-id="05300-152">Name</span></span>| <span data-ttu-id="05300-153">类型</span><span class="sxs-lookup"><span data-stu-id="05300-153">Type</span></span>| <span data-ttu-id="05300-154">说明</span><span class="sxs-lookup"><span data-stu-id="05300-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="05300-155">String</span><span class="sxs-lookup"><span data-stu-id="05300-155">String</span></span>|<span data-ttu-id="05300-156">调用成功。</span><span class="sxs-lookup"><span data-stu-id="05300-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="05300-157">String</span><span class="sxs-lookup"><span data-stu-id="05300-157">String</span></span>|<span data-ttu-id="05300-158">调用失败。</span><span class="sxs-lookup"><span data-stu-id="05300-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="05300-159">要求</span><span class="sxs-lookup"><span data-stu-id="05300-159">Requirements</span></span>

|<span data-ttu-id="05300-160">要求</span><span class="sxs-lookup"><span data-stu-id="05300-160">Requirement</span></span>| <span data-ttu-id="05300-161">值</span><span class="sxs-lookup"><span data-stu-id="05300-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="05300-162">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="05300-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05300-163">1.1</span><span class="sxs-lookup"><span data-stu-id="05300-163">1.1</span></span>|
|[<span data-ttu-id="05300-164">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="05300-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05300-165">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="05300-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="05300-166">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="05300-166">CoercionType: String</span></span>

<span data-ttu-id="05300-167">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="05300-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="05300-168">类型</span><span class="sxs-lookup"><span data-stu-id="05300-168">Type</span></span>

*   <span data-ttu-id="05300-169">String</span><span class="sxs-lookup"><span data-stu-id="05300-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="05300-170">属性：</span><span class="sxs-lookup"><span data-stu-id="05300-170">Properties:</span></span>

|<span data-ttu-id="05300-171">名称</span><span class="sxs-lookup"><span data-stu-id="05300-171">Name</span></span>| <span data-ttu-id="05300-172">类型</span><span class="sxs-lookup"><span data-stu-id="05300-172">Type</span></span>| <span data-ttu-id="05300-173">说明</span><span class="sxs-lookup"><span data-stu-id="05300-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="05300-174">String</span><span class="sxs-lookup"><span data-stu-id="05300-174">String</span></span>|<span data-ttu-id="05300-175">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="05300-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="05300-176">String</span><span class="sxs-lookup"><span data-stu-id="05300-176">String</span></span>|<span data-ttu-id="05300-177">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="05300-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="05300-178">要求</span><span class="sxs-lookup"><span data-stu-id="05300-178">Requirements</span></span>

|<span data-ttu-id="05300-179">要求</span><span class="sxs-lookup"><span data-stu-id="05300-179">Requirement</span></span>| <span data-ttu-id="05300-180">值</span><span class="sxs-lookup"><span data-stu-id="05300-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="05300-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="05300-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05300-182">1.1</span><span class="sxs-lookup"><span data-stu-id="05300-182">1.1</span></span>|
|[<span data-ttu-id="05300-183">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="05300-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05300-184">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="05300-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="05300-185">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="05300-185">SourceProperty: String</span></span>

<span data-ttu-id="05300-186">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="05300-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="05300-187">类型</span><span class="sxs-lookup"><span data-stu-id="05300-187">Type</span></span>

*   <span data-ttu-id="05300-188">String</span><span class="sxs-lookup"><span data-stu-id="05300-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="05300-189">属性：</span><span class="sxs-lookup"><span data-stu-id="05300-189">Properties:</span></span>

|<span data-ttu-id="05300-190">名称</span><span class="sxs-lookup"><span data-stu-id="05300-190">Name</span></span>| <span data-ttu-id="05300-191">类型</span><span class="sxs-lookup"><span data-stu-id="05300-191">Type</span></span>| <span data-ttu-id="05300-192">说明</span><span class="sxs-lookup"><span data-stu-id="05300-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="05300-193">String</span><span class="sxs-lookup"><span data-stu-id="05300-193">String</span></span>|<span data-ttu-id="05300-194">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="05300-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="05300-195">String</span><span class="sxs-lookup"><span data-stu-id="05300-195">String</span></span>|<span data-ttu-id="05300-196">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="05300-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="05300-197">要求</span><span class="sxs-lookup"><span data-stu-id="05300-197">Requirements</span></span>

|<span data-ttu-id="05300-198">要求</span><span class="sxs-lookup"><span data-stu-id="05300-198">Requirement</span></span>| <span data-ttu-id="05300-199">值</span><span class="sxs-lookup"><span data-stu-id="05300-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="05300-200">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="05300-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05300-201">1.1</span><span class="sxs-lookup"><span data-stu-id="05300-201">1.1</span></span>|
|[<span data-ttu-id="05300-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="05300-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05300-203">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="05300-203">Compose or Read</span></span>|
