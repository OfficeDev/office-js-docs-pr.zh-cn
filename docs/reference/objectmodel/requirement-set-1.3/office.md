---
title: Office 命名空间-要求集1。3
description: Outlook 外接程序 API 的顶级命名空间的对象模型（邮箱 API 1.3 版本）。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 706f12f4425a883f0d18fcd6f9ee18972972d72b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717773"
---
# <a name="office"></a><span data-ttu-id="4ddac-103">Office</span><span class="sxs-lookup"><span data-stu-id="4ddac-103">Office</span></span>

<span data-ttu-id="4ddac-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="4ddac-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ddac-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4ddac-106">Requirements</span></span>

|<span data-ttu-id="4ddac-107">要求</span><span class="sxs-lookup"><span data-stu-id="4ddac-107">Requirement</span></span>| <span data-ttu-id="4ddac-108">值</span><span class="sxs-lookup"><span data-stu-id="4ddac-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ddac-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4ddac-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4ddac-110">1.1</span><span class="sxs-lookup"><span data-stu-id="4ddac-110">1.1</span></span>|
|[<span data-ttu-id="4ddac-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4ddac-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4ddac-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4ddac-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="4ddac-113">属性</span><span class="sxs-lookup"><span data-stu-id="4ddac-113">Properties</span></span>

| <span data-ttu-id="4ddac-114">属性</span><span class="sxs-lookup"><span data-stu-id="4ddac-114">Property</span></span> | <span data-ttu-id="4ddac-115">型号</span><span class="sxs-lookup"><span data-stu-id="4ddac-115">Modes</span></span> | <span data-ttu-id="4ddac-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="4ddac-116">Return type</span></span> | <span data-ttu-id="4ddac-117">最低</span><span class="sxs-lookup"><span data-stu-id="4ddac-117">Minimum</span></span><br><span data-ttu-id="4ddac-118">要求集</span><span class="sxs-lookup"><span data-stu-id="4ddac-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4ddac-119">context</span><span class="sxs-lookup"><span data-stu-id="4ddac-119">context</span></span>](office.context.md) | <span data-ttu-id="4ddac-120">撰写</span><span class="sxs-lookup"><span data-stu-id="4ddac-120">Compose</span></span><br><span data-ttu-id="4ddac-121">读取</span><span class="sxs-lookup"><span data-stu-id="4ddac-121">Read</span></span> | [<span data-ttu-id="4ddac-122">Context</span><span class="sxs-lookup"><span data-stu-id="4ddac-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="4ddac-123">1.1</span><span class="sxs-lookup"><span data-stu-id="4ddac-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="4ddac-124">枚举</span><span class="sxs-lookup"><span data-stu-id="4ddac-124">Enumerations</span></span>

| <span data-ttu-id="4ddac-125">枚举</span><span class="sxs-lookup"><span data-stu-id="4ddac-125">Enumeration</span></span> | <span data-ttu-id="4ddac-126">型号</span><span class="sxs-lookup"><span data-stu-id="4ddac-126">Modes</span></span> | <span data-ttu-id="4ddac-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="4ddac-127">Return type</span></span> | <span data-ttu-id="4ddac-128">最低</span><span class="sxs-lookup"><span data-stu-id="4ddac-128">Minimum</span></span><br><span data-ttu-id="4ddac-129">要求集</span><span class="sxs-lookup"><span data-stu-id="4ddac-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4ddac-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4ddac-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4ddac-131">撰写</span><span class="sxs-lookup"><span data-stu-id="4ddac-131">Compose</span></span><br><span data-ttu-id="4ddac-132">读取</span><span class="sxs-lookup"><span data-stu-id="4ddac-132">Read</span></span> | <span data-ttu-id="4ddac-133">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-133">String</span></span> | [<span data-ttu-id="4ddac-134">1.1</span><span class="sxs-lookup"><span data-stu-id="4ddac-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4ddac-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4ddac-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4ddac-136">撰写</span><span class="sxs-lookup"><span data-stu-id="4ddac-136">Compose</span></span><br><span data-ttu-id="4ddac-137">读取</span><span class="sxs-lookup"><span data-stu-id="4ddac-137">Read</span></span> | <span data-ttu-id="4ddac-138">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-138">String</span></span> | [<span data-ttu-id="4ddac-139">1.1</span><span class="sxs-lookup"><span data-stu-id="4ddac-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4ddac-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4ddac-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4ddac-141">撰写</span><span class="sxs-lookup"><span data-stu-id="4ddac-141">Compose</span></span><br><span data-ttu-id="4ddac-142">读取</span><span class="sxs-lookup"><span data-stu-id="4ddac-142">Read</span></span> | <span data-ttu-id="4ddac-143">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-143">String</span></span> | [<span data-ttu-id="4ddac-144">1.1</span><span class="sxs-lookup"><span data-stu-id="4ddac-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="4ddac-145">命名空间</span><span class="sxs-lookup"><span data-stu-id="4ddac-145">Namespaces</span></span>

<span data-ttu-id="4ddac-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="4ddac-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="4ddac-147">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="4ddac-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4ddac-148">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="4ddac-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="4ddac-149">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="4ddac-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4ddac-150">类型</span><span class="sxs-lookup"><span data-stu-id="4ddac-150">Type</span></span>

*   <span data-ttu-id="4ddac-151">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ddac-152">属性：</span><span class="sxs-lookup"><span data-stu-id="4ddac-152">Properties:</span></span>

|<span data-ttu-id="4ddac-153">姓名</span><span class="sxs-lookup"><span data-stu-id="4ddac-153">Name</span></span>| <span data-ttu-id="4ddac-154">类型</span><span class="sxs-lookup"><span data-stu-id="4ddac-154">Type</span></span>| <span data-ttu-id="4ddac-155">说明</span><span class="sxs-lookup"><span data-stu-id="4ddac-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4ddac-156">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-156">String</span></span>|<span data-ttu-id="4ddac-157">调用成功。</span><span class="sxs-lookup"><span data-stu-id="4ddac-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4ddac-158">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-158">String</span></span>|<span data-ttu-id="4ddac-159">调用失败。</span><span class="sxs-lookup"><span data-stu-id="4ddac-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ddac-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="4ddac-160">Requirements</span></span>

|<span data-ttu-id="4ddac-161">要求</span><span class="sxs-lookup"><span data-stu-id="4ddac-161">Requirement</span></span>| <span data-ttu-id="4ddac-162">值</span><span class="sxs-lookup"><span data-stu-id="4ddac-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ddac-163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4ddac-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4ddac-164">1.1</span><span class="sxs-lookup"><span data-stu-id="4ddac-164">1.1</span></span>|
|[<span data-ttu-id="4ddac-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4ddac-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4ddac-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4ddac-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="4ddac-167">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="4ddac-167">CoercionType: String</span></span>

<span data-ttu-id="4ddac-168">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="4ddac-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4ddac-169">类型</span><span class="sxs-lookup"><span data-stu-id="4ddac-169">Type</span></span>

*   <span data-ttu-id="4ddac-170">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ddac-171">属性：</span><span class="sxs-lookup"><span data-stu-id="4ddac-171">Properties:</span></span>

|<span data-ttu-id="4ddac-172">姓名</span><span class="sxs-lookup"><span data-stu-id="4ddac-172">Name</span></span>| <span data-ttu-id="4ddac-173">类型</span><span class="sxs-lookup"><span data-stu-id="4ddac-173">Type</span></span>| <span data-ttu-id="4ddac-174">说明</span><span class="sxs-lookup"><span data-stu-id="4ddac-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4ddac-175">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-175">String</span></span>|<span data-ttu-id="4ddac-176">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="4ddac-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4ddac-177">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-177">String</span></span>|<span data-ttu-id="4ddac-178">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="4ddac-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ddac-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="4ddac-179">Requirements</span></span>

|<span data-ttu-id="4ddac-180">要求</span><span class="sxs-lookup"><span data-stu-id="4ddac-180">Requirement</span></span>| <span data-ttu-id="4ddac-181">值</span><span class="sxs-lookup"><span data-stu-id="4ddac-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ddac-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4ddac-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4ddac-183">1.1</span><span class="sxs-lookup"><span data-stu-id="4ddac-183">1.1</span></span>|
|[<span data-ttu-id="4ddac-184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4ddac-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4ddac-185">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4ddac-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="4ddac-186">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="4ddac-186">SourceProperty: String</span></span>

<span data-ttu-id="4ddac-187">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="4ddac-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4ddac-188">类型</span><span class="sxs-lookup"><span data-stu-id="4ddac-188">Type</span></span>

*   <span data-ttu-id="4ddac-189">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ddac-190">属性：</span><span class="sxs-lookup"><span data-stu-id="4ddac-190">Properties:</span></span>

|<span data-ttu-id="4ddac-191">姓名</span><span class="sxs-lookup"><span data-stu-id="4ddac-191">Name</span></span>| <span data-ttu-id="4ddac-192">类型</span><span class="sxs-lookup"><span data-stu-id="4ddac-192">Type</span></span>| <span data-ttu-id="4ddac-193">说明</span><span class="sxs-lookup"><span data-stu-id="4ddac-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4ddac-194">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-194">String</span></span>|<span data-ttu-id="4ddac-195">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="4ddac-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4ddac-196">String</span><span class="sxs-lookup"><span data-stu-id="4ddac-196">String</span></span>|<span data-ttu-id="4ddac-197">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="4ddac-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ddac-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="4ddac-198">Requirements</span></span>

|<span data-ttu-id="4ddac-199">要求</span><span class="sxs-lookup"><span data-stu-id="4ddac-199">Requirement</span></span>| <span data-ttu-id="4ddac-200">值</span><span class="sxs-lookup"><span data-stu-id="4ddac-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ddac-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4ddac-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4ddac-202">1.1</span><span class="sxs-lookup"><span data-stu-id="4ddac-202">1.1</span></span>|
|[<span data-ttu-id="4ddac-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4ddac-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4ddac-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4ddac-204">Compose or Read</span></span>|
