---
title: Office 命名空间-要求集1。1
description: 使用邮箱 API 要求集1.1 的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: b22dbae0824c8b5047ce90c255e06f09744dd05c
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890758"
---
# <a name="office-mailbox-requirement-set-11"></a><span data-ttu-id="b8f0d-103">Office （邮箱要求集1.1）</span><span class="sxs-lookup"><span data-stu-id="b8f0d-103">Office (Mailbox requirement set 1.1)</span></span>

<span data-ttu-id="b8f0d-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8f0d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="b8f0d-106">Requirements</span></span>

|<span data-ttu-id="b8f0d-107">要求</span><span class="sxs-lookup"><span data-stu-id="b8f0d-107">Requirement</span></span>| <span data-ttu-id="b8f0d-108">值</span><span class="sxs-lookup"><span data-stu-id="b8f0d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8f0d-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8f0d-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b8f0d-110">1.1</span><span class="sxs-lookup"><span data-stu-id="b8f0d-110">1.1</span></span>|
|[<span data-ttu-id="b8f0d-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8f0d-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b8f0d-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8f0d-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b8f0d-113">属性</span><span class="sxs-lookup"><span data-stu-id="b8f0d-113">Properties</span></span>

| <span data-ttu-id="b8f0d-114">属性</span><span class="sxs-lookup"><span data-stu-id="b8f0d-114">Property</span></span> | <span data-ttu-id="b8f0d-115">型号</span><span class="sxs-lookup"><span data-stu-id="b8f0d-115">Modes</span></span> | <span data-ttu-id="b8f0d-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="b8f0d-116">Return type</span></span> | <span data-ttu-id="b8f0d-117">最低</span><span class="sxs-lookup"><span data-stu-id="b8f0d-117">Minimum</span></span><br><span data-ttu-id="b8f0d-118">要求集</span><span class="sxs-lookup"><span data-stu-id="b8f0d-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b8f0d-119">context</span><span class="sxs-lookup"><span data-stu-id="b8f0d-119">context</span></span>](office.context.md) | <span data-ttu-id="b8f0d-120">撰写</span><span class="sxs-lookup"><span data-stu-id="b8f0d-120">Compose</span></span><br><span data-ttu-id="b8f0d-121">读取</span><span class="sxs-lookup"><span data-stu-id="b8f0d-121">Read</span></span> | [<span data-ttu-id="b8f0d-122">Context</span><span class="sxs-lookup"><span data-stu-id="b8f0d-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="b8f0d-123">1.1</span><span class="sxs-lookup"><span data-stu-id="b8f0d-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="b8f0d-124">枚举</span><span class="sxs-lookup"><span data-stu-id="b8f0d-124">Enumerations</span></span>

| <span data-ttu-id="b8f0d-125">枚举</span><span class="sxs-lookup"><span data-stu-id="b8f0d-125">Enumeration</span></span> | <span data-ttu-id="b8f0d-126">型号</span><span class="sxs-lookup"><span data-stu-id="b8f0d-126">Modes</span></span> | <span data-ttu-id="b8f0d-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="b8f0d-127">Return type</span></span> | <span data-ttu-id="b8f0d-128">最低</span><span class="sxs-lookup"><span data-stu-id="b8f0d-128">Minimum</span></span><br><span data-ttu-id="b8f0d-129">要求集</span><span class="sxs-lookup"><span data-stu-id="b8f0d-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b8f0d-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b8f0d-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b8f0d-131">撰写</span><span class="sxs-lookup"><span data-stu-id="b8f0d-131">Compose</span></span><br><span data-ttu-id="b8f0d-132">读取</span><span class="sxs-lookup"><span data-stu-id="b8f0d-132">Read</span></span> | <span data-ttu-id="b8f0d-133">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-133">String</span></span> | [<span data-ttu-id="b8f0d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b8f0d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b8f0d-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b8f0d-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b8f0d-136">撰写</span><span class="sxs-lookup"><span data-stu-id="b8f0d-136">Compose</span></span><br><span data-ttu-id="b8f0d-137">读取</span><span class="sxs-lookup"><span data-stu-id="b8f0d-137">Read</span></span> | <span data-ttu-id="b8f0d-138">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-138">String</span></span> | [<span data-ttu-id="b8f0d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b8f0d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b8f0d-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b8f0d-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b8f0d-141">撰写</span><span class="sxs-lookup"><span data-stu-id="b8f0d-141">Compose</span></span><br><span data-ttu-id="b8f0d-142">读取</span><span class="sxs-lookup"><span data-stu-id="b8f0d-142">Read</span></span> | <span data-ttu-id="b8f0d-143">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-143">String</span></span> | [<span data-ttu-id="b8f0d-144">1.1</span><span class="sxs-lookup"><span data-stu-id="b8f0d-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="b8f0d-145">命名空间</span><span class="sxs-lookup"><span data-stu-id="b8f0d-145">Namespaces</span></span>

<span data-ttu-id="b8f0d-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="b8f0d-147">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="b8f0d-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="b8f0d-148">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="b8f0d-149">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b8f0d-150">类型</span><span class="sxs-lookup"><span data-stu-id="b8f0d-150">Type</span></span>

*   <span data-ttu-id="b8f0d-151">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b8f0d-152">属性：</span><span class="sxs-lookup"><span data-stu-id="b8f0d-152">Properties:</span></span>

|<span data-ttu-id="b8f0d-153">姓名</span><span class="sxs-lookup"><span data-stu-id="b8f0d-153">Name</span></span>| <span data-ttu-id="b8f0d-154">类型</span><span class="sxs-lookup"><span data-stu-id="b8f0d-154">Type</span></span>| <span data-ttu-id="b8f0d-155">说明</span><span class="sxs-lookup"><span data-stu-id="b8f0d-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b8f0d-156">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-156">String</span></span>|<span data-ttu-id="b8f0d-157">调用成功。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b8f0d-158">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-158">String</span></span>|<span data-ttu-id="b8f0d-159">调用失败。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b8f0d-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="b8f0d-160">Requirements</span></span>

|<span data-ttu-id="b8f0d-161">要求</span><span class="sxs-lookup"><span data-stu-id="b8f0d-161">Requirement</span></span>| <span data-ttu-id="b8f0d-162">值</span><span class="sxs-lookup"><span data-stu-id="b8f0d-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8f0d-163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8f0d-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b8f0d-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b8f0d-164">1.1</span></span>|
|[<span data-ttu-id="b8f0d-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8f0d-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b8f0d-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8f0d-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="b8f0d-167">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-167">CoercionType: String</span></span>

<span data-ttu-id="b8f0d-168">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b8f0d-169">类型</span><span class="sxs-lookup"><span data-stu-id="b8f0d-169">Type</span></span>

*   <span data-ttu-id="b8f0d-170">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b8f0d-171">属性：</span><span class="sxs-lookup"><span data-stu-id="b8f0d-171">Properties:</span></span>

|<span data-ttu-id="b8f0d-172">姓名</span><span class="sxs-lookup"><span data-stu-id="b8f0d-172">Name</span></span>| <span data-ttu-id="b8f0d-173">类型</span><span class="sxs-lookup"><span data-stu-id="b8f0d-173">Type</span></span>| <span data-ttu-id="b8f0d-174">说明</span><span class="sxs-lookup"><span data-stu-id="b8f0d-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b8f0d-175">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-175">String</span></span>|<span data-ttu-id="b8f0d-176">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b8f0d-177">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-177">String</span></span>|<span data-ttu-id="b8f0d-178">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b8f0d-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="b8f0d-179">Requirements</span></span>

|<span data-ttu-id="b8f0d-180">要求</span><span class="sxs-lookup"><span data-stu-id="b8f0d-180">Requirement</span></span>| <span data-ttu-id="b8f0d-181">值</span><span class="sxs-lookup"><span data-stu-id="b8f0d-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8f0d-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8f0d-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b8f0d-183">1.1</span><span class="sxs-lookup"><span data-stu-id="b8f0d-183">1.1</span></span>|
|[<span data-ttu-id="b8f0d-184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8f0d-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b8f0d-185">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8f0d-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="b8f0d-186">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-186">SourceProperty: String</span></span>

<span data-ttu-id="b8f0d-187">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b8f0d-188">类型</span><span class="sxs-lookup"><span data-stu-id="b8f0d-188">Type</span></span>

*   <span data-ttu-id="b8f0d-189">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b8f0d-190">属性：</span><span class="sxs-lookup"><span data-stu-id="b8f0d-190">Properties:</span></span>

|<span data-ttu-id="b8f0d-191">姓名</span><span class="sxs-lookup"><span data-stu-id="b8f0d-191">Name</span></span>| <span data-ttu-id="b8f0d-192">类型</span><span class="sxs-lookup"><span data-stu-id="b8f0d-192">Type</span></span>| <span data-ttu-id="b8f0d-193">说明</span><span class="sxs-lookup"><span data-stu-id="b8f0d-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b8f0d-194">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-194">String</span></span>|<span data-ttu-id="b8f0d-195">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b8f0d-196">String</span><span class="sxs-lookup"><span data-stu-id="b8f0d-196">String</span></span>|<span data-ttu-id="b8f0d-197">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="b8f0d-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b8f0d-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="b8f0d-198">Requirements</span></span>

|<span data-ttu-id="b8f0d-199">要求</span><span class="sxs-lookup"><span data-stu-id="b8f0d-199">Requirement</span></span>| <span data-ttu-id="b8f0d-200">值</span><span class="sxs-lookup"><span data-stu-id="b8f0d-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8f0d-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b8f0d-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b8f0d-202">1.1</span><span class="sxs-lookup"><span data-stu-id="b8f0d-202">1.1</span></span>|
|[<span data-ttu-id="b8f0d-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b8f0d-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b8f0d-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b8f0d-204">Compose or Read</span></span>|