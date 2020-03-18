---
title: Office 命名空间-要求集1。1
description: Outlook 外接程序 API 的顶级命名空间的对象模型（邮箱 API 1.1 版本）。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e881f0f9eac054f2b95436504da24cc7d4dec86d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720188"
---
# <a name="office"></a><span data-ttu-id="0b6d0-103">Office</span><span class="sxs-lookup"><span data-stu-id="0b6d0-103">Office</span></span>

<span data-ttu-id="0b6d0-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b6d0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="0b6d0-106">Requirements</span></span>

|<span data-ttu-id="0b6d0-107">要求</span><span class="sxs-lookup"><span data-stu-id="0b6d0-107">Requirement</span></span>| <span data-ttu-id="0b6d0-108">值</span><span class="sxs-lookup"><span data-stu-id="0b6d0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b6d0-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0b6d0-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0b6d0-110">1.1</span><span class="sxs-lookup"><span data-stu-id="0b6d0-110">1.1</span></span>|
|[<span data-ttu-id="0b6d0-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0b6d0-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0b6d0-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0b6d0-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="0b6d0-113">属性</span><span class="sxs-lookup"><span data-stu-id="0b6d0-113">Properties</span></span>

| <span data-ttu-id="0b6d0-114">属性</span><span class="sxs-lookup"><span data-stu-id="0b6d0-114">Property</span></span> | <span data-ttu-id="0b6d0-115">型号</span><span class="sxs-lookup"><span data-stu-id="0b6d0-115">Modes</span></span> | <span data-ttu-id="0b6d0-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="0b6d0-116">Return type</span></span> | <span data-ttu-id="0b6d0-117">最低</span><span class="sxs-lookup"><span data-stu-id="0b6d0-117">Minimum</span></span><br><span data-ttu-id="0b6d0-118">要求集</span><span class="sxs-lookup"><span data-stu-id="0b6d0-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0b6d0-119">context</span><span class="sxs-lookup"><span data-stu-id="0b6d0-119">context</span></span>](office.context.md) | <span data-ttu-id="0b6d0-120">撰写</span><span class="sxs-lookup"><span data-stu-id="0b6d0-120">Compose</span></span><br><span data-ttu-id="0b6d0-121">读取</span><span class="sxs-lookup"><span data-stu-id="0b6d0-121">Read</span></span> | [<span data-ttu-id="0b6d0-122">Context</span><span class="sxs-lookup"><span data-stu-id="0b6d0-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="0b6d0-123">1.1</span><span class="sxs-lookup"><span data-stu-id="0b6d0-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="0b6d0-124">枚举</span><span class="sxs-lookup"><span data-stu-id="0b6d0-124">Enumerations</span></span>

| <span data-ttu-id="0b6d0-125">枚举</span><span class="sxs-lookup"><span data-stu-id="0b6d0-125">Enumeration</span></span> | <span data-ttu-id="0b6d0-126">型号</span><span class="sxs-lookup"><span data-stu-id="0b6d0-126">Modes</span></span> | <span data-ttu-id="0b6d0-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="0b6d0-127">Return type</span></span> | <span data-ttu-id="0b6d0-128">最低</span><span class="sxs-lookup"><span data-stu-id="0b6d0-128">Minimum</span></span><br><span data-ttu-id="0b6d0-129">要求集</span><span class="sxs-lookup"><span data-stu-id="0b6d0-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0b6d0-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0b6d0-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0b6d0-131">撰写</span><span class="sxs-lookup"><span data-stu-id="0b6d0-131">Compose</span></span><br><span data-ttu-id="0b6d0-132">读取</span><span class="sxs-lookup"><span data-stu-id="0b6d0-132">Read</span></span> | <span data-ttu-id="0b6d0-133">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-133">String</span></span> | [<span data-ttu-id="0b6d0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="0b6d0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0b6d0-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0b6d0-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0b6d0-136">撰写</span><span class="sxs-lookup"><span data-stu-id="0b6d0-136">Compose</span></span><br><span data-ttu-id="0b6d0-137">读取</span><span class="sxs-lookup"><span data-stu-id="0b6d0-137">Read</span></span> | <span data-ttu-id="0b6d0-138">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-138">String</span></span> | [<span data-ttu-id="0b6d0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="0b6d0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0b6d0-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0b6d0-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0b6d0-141">撰写</span><span class="sxs-lookup"><span data-stu-id="0b6d0-141">Compose</span></span><br><span data-ttu-id="0b6d0-142">读取</span><span class="sxs-lookup"><span data-stu-id="0b6d0-142">Read</span></span> | <span data-ttu-id="0b6d0-143">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-143">String</span></span> | [<span data-ttu-id="0b6d0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="0b6d0-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="0b6d0-145">命名空间</span><span class="sxs-lookup"><span data-stu-id="0b6d0-145">Namespaces</span></span>

<span data-ttu-id="0b6d0-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="0b6d0-147">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="0b6d0-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="0b6d0-148">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="0b6d0-149">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0b6d0-150">类型</span><span class="sxs-lookup"><span data-stu-id="0b6d0-150">Type</span></span>

*   <span data-ttu-id="0b6d0-151">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0b6d0-152">属性：</span><span class="sxs-lookup"><span data-stu-id="0b6d0-152">Properties:</span></span>

|<span data-ttu-id="0b6d0-153">姓名</span><span class="sxs-lookup"><span data-stu-id="0b6d0-153">Name</span></span>| <span data-ttu-id="0b6d0-154">类型</span><span class="sxs-lookup"><span data-stu-id="0b6d0-154">Type</span></span>| <span data-ttu-id="0b6d0-155">说明</span><span class="sxs-lookup"><span data-stu-id="0b6d0-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0b6d0-156">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-156">String</span></span>|<span data-ttu-id="0b6d0-157">调用成功。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0b6d0-158">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-158">String</span></span>|<span data-ttu-id="0b6d0-159">调用失败。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b6d0-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="0b6d0-160">Requirements</span></span>

|<span data-ttu-id="0b6d0-161">要求</span><span class="sxs-lookup"><span data-stu-id="0b6d0-161">Requirement</span></span>| <span data-ttu-id="0b6d0-162">值</span><span class="sxs-lookup"><span data-stu-id="0b6d0-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b6d0-163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0b6d0-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0b6d0-164">1.1</span><span class="sxs-lookup"><span data-stu-id="0b6d0-164">1.1</span></span>|
|[<span data-ttu-id="0b6d0-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0b6d0-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0b6d0-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0b6d0-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="0b6d0-167">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-167">CoercionType: String</span></span>

<span data-ttu-id="0b6d0-168">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0b6d0-169">类型</span><span class="sxs-lookup"><span data-stu-id="0b6d0-169">Type</span></span>

*   <span data-ttu-id="0b6d0-170">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0b6d0-171">属性：</span><span class="sxs-lookup"><span data-stu-id="0b6d0-171">Properties:</span></span>

|<span data-ttu-id="0b6d0-172">姓名</span><span class="sxs-lookup"><span data-stu-id="0b6d0-172">Name</span></span>| <span data-ttu-id="0b6d0-173">类型</span><span class="sxs-lookup"><span data-stu-id="0b6d0-173">Type</span></span>| <span data-ttu-id="0b6d0-174">说明</span><span class="sxs-lookup"><span data-stu-id="0b6d0-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0b6d0-175">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-175">String</span></span>|<span data-ttu-id="0b6d0-176">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0b6d0-177">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-177">String</span></span>|<span data-ttu-id="0b6d0-178">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b6d0-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="0b6d0-179">Requirements</span></span>

|<span data-ttu-id="0b6d0-180">要求</span><span class="sxs-lookup"><span data-stu-id="0b6d0-180">Requirement</span></span>| <span data-ttu-id="0b6d0-181">值</span><span class="sxs-lookup"><span data-stu-id="0b6d0-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b6d0-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0b6d0-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0b6d0-183">1.1</span><span class="sxs-lookup"><span data-stu-id="0b6d0-183">1.1</span></span>|
|[<span data-ttu-id="0b6d0-184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0b6d0-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0b6d0-185">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0b6d0-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="0b6d0-186">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-186">SourceProperty: String</span></span>

<span data-ttu-id="0b6d0-187">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0b6d0-188">类型</span><span class="sxs-lookup"><span data-stu-id="0b6d0-188">Type</span></span>

*   <span data-ttu-id="0b6d0-189">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0b6d0-190">属性：</span><span class="sxs-lookup"><span data-stu-id="0b6d0-190">Properties:</span></span>

|<span data-ttu-id="0b6d0-191">姓名</span><span class="sxs-lookup"><span data-stu-id="0b6d0-191">Name</span></span>| <span data-ttu-id="0b6d0-192">类型</span><span class="sxs-lookup"><span data-stu-id="0b6d0-192">Type</span></span>| <span data-ttu-id="0b6d0-193">说明</span><span class="sxs-lookup"><span data-stu-id="0b6d0-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0b6d0-194">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-194">String</span></span>|<span data-ttu-id="0b6d0-195">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0b6d0-196">String</span><span class="sxs-lookup"><span data-stu-id="0b6d0-196">String</span></span>|<span data-ttu-id="0b6d0-197">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="0b6d0-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b6d0-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="0b6d0-198">Requirements</span></span>

|<span data-ttu-id="0b6d0-199">要求</span><span class="sxs-lookup"><span data-stu-id="0b6d0-199">Requirement</span></span>| <span data-ttu-id="0b6d0-200">值</span><span class="sxs-lookup"><span data-stu-id="0b6d0-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b6d0-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0b6d0-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0b6d0-202">1.1</span><span class="sxs-lookup"><span data-stu-id="0b6d0-202">1.1</span></span>|
|[<span data-ttu-id="0b6d0-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0b6d0-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0b6d0-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0b6d0-204">Compose or Read</span></span>|