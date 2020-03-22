---
title: Office 命名空间-要求集1。3
description: 使用邮箱 API 要求集1.3 的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: eb3ebba75df8345402ab0ce4ac2b5cc5f0354e6c
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890695"
---
# <a name="office-mailbox-requirement-set-13"></a><span data-ttu-id="b69db-103">Office （邮箱要求集1.3）</span><span class="sxs-lookup"><span data-stu-id="b69db-103">Office (Mailbox requirement set 1.3)</span></span>

<span data-ttu-id="b69db-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="b69db-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b69db-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="b69db-106">Requirements</span></span>

|<span data-ttu-id="b69db-107">要求</span><span class="sxs-lookup"><span data-stu-id="b69db-107">Requirement</span></span>| <span data-ttu-id="b69db-108">值</span><span class="sxs-lookup"><span data-stu-id="b69db-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b69db-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b69db-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b69db-110">1.1</span><span class="sxs-lookup"><span data-stu-id="b69db-110">1.1</span></span>|
|[<span data-ttu-id="b69db-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b69db-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b69db-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b69db-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b69db-113">属性</span><span class="sxs-lookup"><span data-stu-id="b69db-113">Properties</span></span>

| <span data-ttu-id="b69db-114">属性</span><span class="sxs-lookup"><span data-stu-id="b69db-114">Property</span></span> | <span data-ttu-id="b69db-115">型号</span><span class="sxs-lookup"><span data-stu-id="b69db-115">Modes</span></span> | <span data-ttu-id="b69db-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="b69db-116">Return type</span></span> | <span data-ttu-id="b69db-117">最低</span><span class="sxs-lookup"><span data-stu-id="b69db-117">Minimum</span></span><br><span data-ttu-id="b69db-118">要求集</span><span class="sxs-lookup"><span data-stu-id="b69db-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b69db-119">context</span><span class="sxs-lookup"><span data-stu-id="b69db-119">context</span></span>](office.context.md) | <span data-ttu-id="b69db-120">撰写</span><span class="sxs-lookup"><span data-stu-id="b69db-120">Compose</span></span><br><span data-ttu-id="b69db-121">读取</span><span class="sxs-lookup"><span data-stu-id="b69db-121">Read</span></span> | [<span data-ttu-id="b69db-122">Context</span><span class="sxs-lookup"><span data-stu-id="b69db-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="b69db-123">1.1</span><span class="sxs-lookup"><span data-stu-id="b69db-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="b69db-124">枚举</span><span class="sxs-lookup"><span data-stu-id="b69db-124">Enumerations</span></span>

| <span data-ttu-id="b69db-125">枚举</span><span class="sxs-lookup"><span data-stu-id="b69db-125">Enumeration</span></span> | <span data-ttu-id="b69db-126">型号</span><span class="sxs-lookup"><span data-stu-id="b69db-126">Modes</span></span> | <span data-ttu-id="b69db-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="b69db-127">Return type</span></span> | <span data-ttu-id="b69db-128">最低</span><span class="sxs-lookup"><span data-stu-id="b69db-128">Minimum</span></span><br><span data-ttu-id="b69db-129">要求集</span><span class="sxs-lookup"><span data-stu-id="b69db-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b69db-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b69db-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b69db-131">撰写</span><span class="sxs-lookup"><span data-stu-id="b69db-131">Compose</span></span><br><span data-ttu-id="b69db-132">读取</span><span class="sxs-lookup"><span data-stu-id="b69db-132">Read</span></span> | <span data-ttu-id="b69db-133">String</span><span class="sxs-lookup"><span data-stu-id="b69db-133">String</span></span> | [<span data-ttu-id="b69db-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b69db-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b69db-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b69db-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b69db-136">撰写</span><span class="sxs-lookup"><span data-stu-id="b69db-136">Compose</span></span><br><span data-ttu-id="b69db-137">读取</span><span class="sxs-lookup"><span data-stu-id="b69db-137">Read</span></span> | <span data-ttu-id="b69db-138">String</span><span class="sxs-lookup"><span data-stu-id="b69db-138">String</span></span> | [<span data-ttu-id="b69db-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b69db-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b69db-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b69db-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b69db-141">撰写</span><span class="sxs-lookup"><span data-stu-id="b69db-141">Compose</span></span><br><span data-ttu-id="b69db-142">读取</span><span class="sxs-lookup"><span data-stu-id="b69db-142">Read</span></span> | <span data-ttu-id="b69db-143">String</span><span class="sxs-lookup"><span data-stu-id="b69db-143">String</span></span> | [<span data-ttu-id="b69db-144">1.1</span><span class="sxs-lookup"><span data-stu-id="b69db-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="b69db-145">命名空间</span><span class="sxs-lookup"><span data-stu-id="b69db-145">Namespaces</span></span>

<span data-ttu-id="b69db-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="b69db-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="b69db-147">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="b69db-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="b69db-148">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="b69db-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="b69db-149">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="b69db-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b69db-150">类型</span><span class="sxs-lookup"><span data-stu-id="b69db-150">Type</span></span>

*   <span data-ttu-id="b69db-151">String</span><span class="sxs-lookup"><span data-stu-id="b69db-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b69db-152">属性：</span><span class="sxs-lookup"><span data-stu-id="b69db-152">Properties:</span></span>

|<span data-ttu-id="b69db-153">姓名</span><span class="sxs-lookup"><span data-stu-id="b69db-153">Name</span></span>| <span data-ttu-id="b69db-154">类型</span><span class="sxs-lookup"><span data-stu-id="b69db-154">Type</span></span>| <span data-ttu-id="b69db-155">说明</span><span class="sxs-lookup"><span data-stu-id="b69db-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b69db-156">String</span><span class="sxs-lookup"><span data-stu-id="b69db-156">String</span></span>|<span data-ttu-id="b69db-157">调用成功。</span><span class="sxs-lookup"><span data-stu-id="b69db-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b69db-158">String</span><span class="sxs-lookup"><span data-stu-id="b69db-158">String</span></span>|<span data-ttu-id="b69db-159">调用失败。</span><span class="sxs-lookup"><span data-stu-id="b69db-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b69db-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="b69db-160">Requirements</span></span>

|<span data-ttu-id="b69db-161">要求</span><span class="sxs-lookup"><span data-stu-id="b69db-161">Requirement</span></span>| <span data-ttu-id="b69db-162">值</span><span class="sxs-lookup"><span data-stu-id="b69db-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="b69db-163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b69db-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b69db-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b69db-164">1.1</span></span>|
|[<span data-ttu-id="b69db-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b69db-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b69db-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b69db-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="b69db-167">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="b69db-167">CoercionType: String</span></span>

<span data-ttu-id="b69db-168">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="b69db-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b69db-169">类型</span><span class="sxs-lookup"><span data-stu-id="b69db-169">Type</span></span>

*   <span data-ttu-id="b69db-170">String</span><span class="sxs-lookup"><span data-stu-id="b69db-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b69db-171">属性：</span><span class="sxs-lookup"><span data-stu-id="b69db-171">Properties:</span></span>

|<span data-ttu-id="b69db-172">姓名</span><span class="sxs-lookup"><span data-stu-id="b69db-172">Name</span></span>| <span data-ttu-id="b69db-173">类型</span><span class="sxs-lookup"><span data-stu-id="b69db-173">Type</span></span>| <span data-ttu-id="b69db-174">说明</span><span class="sxs-lookup"><span data-stu-id="b69db-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b69db-175">String</span><span class="sxs-lookup"><span data-stu-id="b69db-175">String</span></span>|<span data-ttu-id="b69db-176">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="b69db-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b69db-177">String</span><span class="sxs-lookup"><span data-stu-id="b69db-177">String</span></span>|<span data-ttu-id="b69db-178">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="b69db-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b69db-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="b69db-179">Requirements</span></span>

|<span data-ttu-id="b69db-180">要求</span><span class="sxs-lookup"><span data-stu-id="b69db-180">Requirement</span></span>| <span data-ttu-id="b69db-181">值</span><span class="sxs-lookup"><span data-stu-id="b69db-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="b69db-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b69db-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b69db-183">1.1</span><span class="sxs-lookup"><span data-stu-id="b69db-183">1.1</span></span>|
|[<span data-ttu-id="b69db-184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b69db-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b69db-185">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b69db-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="b69db-186">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="b69db-186">SourceProperty: String</span></span>

<span data-ttu-id="b69db-187">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="b69db-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b69db-188">类型</span><span class="sxs-lookup"><span data-stu-id="b69db-188">Type</span></span>

*   <span data-ttu-id="b69db-189">String</span><span class="sxs-lookup"><span data-stu-id="b69db-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b69db-190">属性：</span><span class="sxs-lookup"><span data-stu-id="b69db-190">Properties:</span></span>

|<span data-ttu-id="b69db-191">姓名</span><span class="sxs-lookup"><span data-stu-id="b69db-191">Name</span></span>| <span data-ttu-id="b69db-192">类型</span><span class="sxs-lookup"><span data-stu-id="b69db-192">Type</span></span>| <span data-ttu-id="b69db-193">说明</span><span class="sxs-lookup"><span data-stu-id="b69db-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b69db-194">String</span><span class="sxs-lookup"><span data-stu-id="b69db-194">String</span></span>|<span data-ttu-id="b69db-195">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="b69db-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b69db-196">String</span><span class="sxs-lookup"><span data-stu-id="b69db-196">String</span></span>|<span data-ttu-id="b69db-197">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="b69db-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b69db-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="b69db-198">Requirements</span></span>

|<span data-ttu-id="b69db-199">要求</span><span class="sxs-lookup"><span data-stu-id="b69db-199">Requirement</span></span>| <span data-ttu-id="b69db-200">值</span><span class="sxs-lookup"><span data-stu-id="b69db-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="b69db-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b69db-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b69db-202">1.1</span><span class="sxs-lookup"><span data-stu-id="b69db-202">1.1</span></span>|
|[<span data-ttu-id="b69db-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b69db-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b69db-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b69db-204">Compose or Read</span></span>|
