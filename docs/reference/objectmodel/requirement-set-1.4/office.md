---
title: Office 命名空间-要求集1。4
description: Outlook 外接程序 API 的顶级命名空间的对象模型（邮箱 API 1.4 版本）。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e5a5c6de5bb87cb32968d9d9d80c621f0acc238d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720055"
---
# <a name="office"></a><span data-ttu-id="8b28a-103">Office</span><span class="sxs-lookup"><span data-stu-id="8b28a-103">Office</span></span>

<span data-ttu-id="8b28a-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="8b28a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b28a-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b28a-106">Requirements</span></span>

|<span data-ttu-id="8b28a-107">要求</span><span class="sxs-lookup"><span data-stu-id="8b28a-107">Requirement</span></span>| <span data-ttu-id="8b28a-108">值</span><span class="sxs-lookup"><span data-stu-id="8b28a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b28a-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8b28a-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b28a-110">1.1</span><span class="sxs-lookup"><span data-stu-id="8b28a-110">1.1</span></span>|
|[<span data-ttu-id="8b28a-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8b28a-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b28a-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8b28a-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8b28a-113">属性</span><span class="sxs-lookup"><span data-stu-id="8b28a-113">Properties</span></span>

| <span data-ttu-id="8b28a-114">属性</span><span class="sxs-lookup"><span data-stu-id="8b28a-114">Property</span></span> | <span data-ttu-id="8b28a-115">型号</span><span class="sxs-lookup"><span data-stu-id="8b28a-115">Modes</span></span> | <span data-ttu-id="8b28a-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="8b28a-116">Return type</span></span> | <span data-ttu-id="8b28a-117">最低</span><span class="sxs-lookup"><span data-stu-id="8b28a-117">Minimum</span></span><br><span data-ttu-id="8b28a-118">要求集</span><span class="sxs-lookup"><span data-stu-id="8b28a-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8b28a-119">context</span><span class="sxs-lookup"><span data-stu-id="8b28a-119">context</span></span>](office.context.md) | <span data-ttu-id="8b28a-120">撰写</span><span class="sxs-lookup"><span data-stu-id="8b28a-120">Compose</span></span><br><span data-ttu-id="8b28a-121">读取</span><span class="sxs-lookup"><span data-stu-id="8b28a-121">Read</span></span> | [<span data-ttu-id="8b28a-122">Context</span><span class="sxs-lookup"><span data-stu-id="8b28a-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4) | [<span data-ttu-id="8b28a-123">1.1</span><span class="sxs-lookup"><span data-stu-id="8b28a-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="8b28a-124">枚举</span><span class="sxs-lookup"><span data-stu-id="8b28a-124">Enumerations</span></span>

| <span data-ttu-id="8b28a-125">枚举</span><span class="sxs-lookup"><span data-stu-id="8b28a-125">Enumeration</span></span> | <span data-ttu-id="8b28a-126">型号</span><span class="sxs-lookup"><span data-stu-id="8b28a-126">Modes</span></span> | <span data-ttu-id="8b28a-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="8b28a-127">Return type</span></span> | <span data-ttu-id="8b28a-128">最低</span><span class="sxs-lookup"><span data-stu-id="8b28a-128">Minimum</span></span><br><span data-ttu-id="8b28a-129">要求集</span><span class="sxs-lookup"><span data-stu-id="8b28a-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8b28a-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8b28a-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8b28a-131">撰写</span><span class="sxs-lookup"><span data-stu-id="8b28a-131">Compose</span></span><br><span data-ttu-id="8b28a-132">读取</span><span class="sxs-lookup"><span data-stu-id="8b28a-132">Read</span></span> | <span data-ttu-id="8b28a-133">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-133">String</span></span> | [<span data-ttu-id="8b28a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8b28a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b28a-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8b28a-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8b28a-136">撰写</span><span class="sxs-lookup"><span data-stu-id="8b28a-136">Compose</span></span><br><span data-ttu-id="8b28a-137">读取</span><span class="sxs-lookup"><span data-stu-id="8b28a-137">Read</span></span> | <span data-ttu-id="8b28a-138">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-138">String</span></span> | [<span data-ttu-id="8b28a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8b28a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b28a-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8b28a-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8b28a-141">撰写</span><span class="sxs-lookup"><span data-stu-id="8b28a-141">Compose</span></span><br><span data-ttu-id="8b28a-142">读取</span><span class="sxs-lookup"><span data-stu-id="8b28a-142">Read</span></span> | <span data-ttu-id="8b28a-143">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-143">String</span></span> | [<span data-ttu-id="8b28a-144">1.1</span><span class="sxs-lookup"><span data-stu-id="8b28a-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="8b28a-145">命名空间</span><span class="sxs-lookup"><span data-stu-id="8b28a-145">Namespaces</span></span>

<span data-ttu-id="8b28a-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="8b28a-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="8b28a-147">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="8b28a-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="8b28a-148">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="8b28a-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="8b28a-149">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="8b28a-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8b28a-150">类型</span><span class="sxs-lookup"><span data-stu-id="8b28a-150">Type</span></span>

*   <span data-ttu-id="8b28a-151">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8b28a-152">属性：</span><span class="sxs-lookup"><span data-stu-id="8b28a-152">Properties:</span></span>

|<span data-ttu-id="8b28a-153">姓名</span><span class="sxs-lookup"><span data-stu-id="8b28a-153">Name</span></span>| <span data-ttu-id="8b28a-154">类型</span><span class="sxs-lookup"><span data-stu-id="8b28a-154">Type</span></span>| <span data-ttu-id="8b28a-155">说明</span><span class="sxs-lookup"><span data-stu-id="8b28a-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8b28a-156">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-156">String</span></span>|<span data-ttu-id="8b28a-157">调用成功。</span><span class="sxs-lookup"><span data-stu-id="8b28a-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8b28a-158">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-158">String</span></span>|<span data-ttu-id="8b28a-159">调用失败。</span><span class="sxs-lookup"><span data-stu-id="8b28a-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b28a-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b28a-160">Requirements</span></span>

|<span data-ttu-id="8b28a-161">要求</span><span class="sxs-lookup"><span data-stu-id="8b28a-161">Requirement</span></span>| <span data-ttu-id="8b28a-162">值</span><span class="sxs-lookup"><span data-stu-id="8b28a-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b28a-163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8b28a-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b28a-164">1.1</span><span class="sxs-lookup"><span data-stu-id="8b28a-164">1.1</span></span>|
|[<span data-ttu-id="8b28a-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8b28a-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b28a-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8b28a-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="8b28a-167">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="8b28a-167">CoercionType: String</span></span>

<span data-ttu-id="8b28a-168">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="8b28a-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8b28a-169">类型</span><span class="sxs-lookup"><span data-stu-id="8b28a-169">Type</span></span>

*   <span data-ttu-id="8b28a-170">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8b28a-171">属性：</span><span class="sxs-lookup"><span data-stu-id="8b28a-171">Properties:</span></span>

|<span data-ttu-id="8b28a-172">姓名</span><span class="sxs-lookup"><span data-stu-id="8b28a-172">Name</span></span>| <span data-ttu-id="8b28a-173">类型</span><span class="sxs-lookup"><span data-stu-id="8b28a-173">Type</span></span>| <span data-ttu-id="8b28a-174">说明</span><span class="sxs-lookup"><span data-stu-id="8b28a-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8b28a-175">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-175">String</span></span>|<span data-ttu-id="8b28a-176">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="8b28a-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8b28a-177">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-177">String</span></span>|<span data-ttu-id="8b28a-178">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="8b28a-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b28a-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b28a-179">Requirements</span></span>

|<span data-ttu-id="8b28a-180">要求</span><span class="sxs-lookup"><span data-stu-id="8b28a-180">Requirement</span></span>| <span data-ttu-id="8b28a-181">值</span><span class="sxs-lookup"><span data-stu-id="8b28a-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b28a-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8b28a-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b28a-183">1.1</span><span class="sxs-lookup"><span data-stu-id="8b28a-183">1.1</span></span>|
|[<span data-ttu-id="8b28a-184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8b28a-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b28a-185">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8b28a-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="8b28a-186">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="8b28a-186">SourceProperty: String</span></span>

<span data-ttu-id="8b28a-187">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="8b28a-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8b28a-188">类型</span><span class="sxs-lookup"><span data-stu-id="8b28a-188">Type</span></span>

*   <span data-ttu-id="8b28a-189">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8b28a-190">属性：</span><span class="sxs-lookup"><span data-stu-id="8b28a-190">Properties:</span></span>

|<span data-ttu-id="8b28a-191">姓名</span><span class="sxs-lookup"><span data-stu-id="8b28a-191">Name</span></span>| <span data-ttu-id="8b28a-192">类型</span><span class="sxs-lookup"><span data-stu-id="8b28a-192">Type</span></span>| <span data-ttu-id="8b28a-193">说明</span><span class="sxs-lookup"><span data-stu-id="8b28a-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8b28a-194">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-194">String</span></span>|<span data-ttu-id="8b28a-195">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="8b28a-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8b28a-196">String</span><span class="sxs-lookup"><span data-stu-id="8b28a-196">String</span></span>|<span data-ttu-id="8b28a-197">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="8b28a-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b28a-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b28a-198">Requirements</span></span>

|<span data-ttu-id="8b28a-199">要求</span><span class="sxs-lookup"><span data-stu-id="8b28a-199">Requirement</span></span>| <span data-ttu-id="8b28a-200">值</span><span class="sxs-lookup"><span data-stu-id="8b28a-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b28a-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8b28a-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b28a-202">1.1</span><span class="sxs-lookup"><span data-stu-id="8b28a-202">1.1</span></span>|
|[<span data-ttu-id="8b28a-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8b28a-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b28a-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8b28a-204">Compose or Read</span></span>|
