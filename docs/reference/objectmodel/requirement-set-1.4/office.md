---
title: Office 命名空间-要求集1。4
description: 使用邮箱 API 要求集1.4 的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f797fe5281d2031a2182249aeb18d740cd114d43
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430741"
---
# <a name="office-mailbox-requirement-set-14"></a><span data-ttu-id="3f089-103">Office (邮箱要求集 1.4) </span><span class="sxs-lookup"><span data-stu-id="3f089-103">Office (Mailbox requirement set 1.4)</span></span>

<span data-ttu-id="3f089-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="3f089-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3f089-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f089-106">Requirements</span></span>

|<span data-ttu-id="3f089-107">要求</span><span class="sxs-lookup"><span data-stu-id="3f089-107">Requirement</span></span>| <span data-ttu-id="3f089-108">值</span><span class="sxs-lookup"><span data-stu-id="3f089-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f089-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3f089-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f089-110">1.1</span><span class="sxs-lookup"><span data-stu-id="3f089-110">1.1</span></span>|
|[<span data-ttu-id="3f089-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3f089-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f089-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3f089-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3f089-113">属性</span><span class="sxs-lookup"><span data-stu-id="3f089-113">Properties</span></span>

| <span data-ttu-id="3f089-114">属性</span><span class="sxs-lookup"><span data-stu-id="3f089-114">Property</span></span> | <span data-ttu-id="3f089-115">型号</span><span class="sxs-lookup"><span data-stu-id="3f089-115">Modes</span></span> | <span data-ttu-id="3f089-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="3f089-116">Return type</span></span> | <span data-ttu-id="3f089-117">最小值</span><span class="sxs-lookup"><span data-stu-id="3f089-117">Minimum</span></span><br><span data-ttu-id="3f089-118">要求集</span><span class="sxs-lookup"><span data-stu-id="3f089-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3f089-119">context</span><span class="sxs-lookup"><span data-stu-id="3f089-119">context</span></span>](office.context.md) | <span data-ttu-id="3f089-120">撰写</span><span class="sxs-lookup"><span data-stu-id="3f089-120">Compose</span></span><br><span data-ttu-id="3f089-121">阅读</span><span class="sxs-lookup"><span data-stu-id="3f089-121">Read</span></span> | [<span data-ttu-id="3f089-122">Context</span><span class="sxs-lookup"><span data-stu-id="3f089-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="3f089-123">1.1</span><span class="sxs-lookup"><span data-stu-id="3f089-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="3f089-124">枚举</span><span class="sxs-lookup"><span data-stu-id="3f089-124">Enumerations</span></span>

| <span data-ttu-id="3f089-125">枚举</span><span class="sxs-lookup"><span data-stu-id="3f089-125">Enumeration</span></span> | <span data-ttu-id="3f089-126">型号</span><span class="sxs-lookup"><span data-stu-id="3f089-126">Modes</span></span> | <span data-ttu-id="3f089-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="3f089-127">Return type</span></span> | <span data-ttu-id="3f089-128">最小值</span><span class="sxs-lookup"><span data-stu-id="3f089-128">Minimum</span></span><br><span data-ttu-id="3f089-129">要求集</span><span class="sxs-lookup"><span data-stu-id="3f089-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3f089-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3f089-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3f089-131">撰写</span><span class="sxs-lookup"><span data-stu-id="3f089-131">Compose</span></span><br><span data-ttu-id="3f089-132">阅读</span><span class="sxs-lookup"><span data-stu-id="3f089-132">Read</span></span> | <span data-ttu-id="3f089-133">String</span><span class="sxs-lookup"><span data-stu-id="3f089-133">String</span></span> | [<span data-ttu-id="3f089-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3f089-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3f089-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3f089-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3f089-136">撰写</span><span class="sxs-lookup"><span data-stu-id="3f089-136">Compose</span></span><br><span data-ttu-id="3f089-137">阅读</span><span class="sxs-lookup"><span data-stu-id="3f089-137">Read</span></span> | <span data-ttu-id="3f089-138">String</span><span class="sxs-lookup"><span data-stu-id="3f089-138">String</span></span> | [<span data-ttu-id="3f089-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3f089-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3f089-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3f089-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3f089-141">撰写</span><span class="sxs-lookup"><span data-stu-id="3f089-141">Compose</span></span><br><span data-ttu-id="3f089-142">阅读</span><span class="sxs-lookup"><span data-stu-id="3f089-142">Read</span></span> | <span data-ttu-id="3f089-143">String</span><span class="sxs-lookup"><span data-stu-id="3f089-143">String</span></span> | [<span data-ttu-id="3f089-144">1.1</span><span class="sxs-lookup"><span data-stu-id="3f089-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="3f089-145">命名空间</span><span class="sxs-lookup"><span data-stu-id="3f089-145">Namespaces</span></span>

<span data-ttu-id="3f089-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true)：包含许多特定于 Outlook 的枚举，例如、、、、、 `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` 和 `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="3f089-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="3f089-147">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="3f089-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3f089-148">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="3f089-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="3f089-149">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="3f089-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3f089-150">类型</span><span class="sxs-lookup"><span data-stu-id="3f089-150">Type</span></span>

*   <span data-ttu-id="3f089-151">String</span><span class="sxs-lookup"><span data-stu-id="3f089-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f089-152">属性：</span><span class="sxs-lookup"><span data-stu-id="3f089-152">Properties:</span></span>

|<span data-ttu-id="3f089-153">名称</span><span class="sxs-lookup"><span data-stu-id="3f089-153">Name</span></span>| <span data-ttu-id="3f089-154">类型</span><span class="sxs-lookup"><span data-stu-id="3f089-154">Type</span></span>| <span data-ttu-id="3f089-155">说明</span><span class="sxs-lookup"><span data-stu-id="3f089-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3f089-156">String</span><span class="sxs-lookup"><span data-stu-id="3f089-156">String</span></span>|<span data-ttu-id="3f089-157">调用成功。</span><span class="sxs-lookup"><span data-stu-id="3f089-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3f089-158">字符串</span><span class="sxs-lookup"><span data-stu-id="3f089-158">String</span></span>|<span data-ttu-id="3f089-159">调用失败。</span><span class="sxs-lookup"><span data-stu-id="3f089-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f089-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f089-160">Requirements</span></span>

|<span data-ttu-id="3f089-161">要求</span><span class="sxs-lookup"><span data-stu-id="3f089-161">Requirement</span></span>| <span data-ttu-id="3f089-162">值</span><span class="sxs-lookup"><span data-stu-id="3f089-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f089-163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3f089-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f089-164">1.1</span><span class="sxs-lookup"><span data-stu-id="3f089-164">1.1</span></span>|
|[<span data-ttu-id="3f089-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3f089-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f089-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3f089-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="3f089-167">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="3f089-167">CoercionType: String</span></span>

<span data-ttu-id="3f089-168">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="3f089-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3f089-169">类型</span><span class="sxs-lookup"><span data-stu-id="3f089-169">Type</span></span>

*   <span data-ttu-id="3f089-170">String</span><span class="sxs-lookup"><span data-stu-id="3f089-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f089-171">属性：</span><span class="sxs-lookup"><span data-stu-id="3f089-171">Properties:</span></span>

|<span data-ttu-id="3f089-172">名称</span><span class="sxs-lookup"><span data-stu-id="3f089-172">Name</span></span>| <span data-ttu-id="3f089-173">类型</span><span class="sxs-lookup"><span data-stu-id="3f089-173">Type</span></span>| <span data-ttu-id="3f089-174">说明</span><span class="sxs-lookup"><span data-stu-id="3f089-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3f089-175">String</span><span class="sxs-lookup"><span data-stu-id="3f089-175">String</span></span>|<span data-ttu-id="3f089-176">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="3f089-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3f089-177">字符串</span><span class="sxs-lookup"><span data-stu-id="3f089-177">String</span></span>|<span data-ttu-id="3f089-178">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="3f089-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f089-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f089-179">Requirements</span></span>

|<span data-ttu-id="3f089-180">要求</span><span class="sxs-lookup"><span data-stu-id="3f089-180">Requirement</span></span>| <span data-ttu-id="3f089-181">值</span><span class="sxs-lookup"><span data-stu-id="3f089-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f089-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3f089-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f089-183">1.1</span><span class="sxs-lookup"><span data-stu-id="3f089-183">1.1</span></span>|
|[<span data-ttu-id="3f089-184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3f089-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f089-185">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3f089-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="3f089-186">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="3f089-186">SourceProperty: String</span></span>

<span data-ttu-id="3f089-187">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="3f089-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3f089-188">类型</span><span class="sxs-lookup"><span data-stu-id="3f089-188">Type</span></span>

*   <span data-ttu-id="3f089-189">String</span><span class="sxs-lookup"><span data-stu-id="3f089-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f089-190">属性：</span><span class="sxs-lookup"><span data-stu-id="3f089-190">Properties:</span></span>

|<span data-ttu-id="3f089-191">名称</span><span class="sxs-lookup"><span data-stu-id="3f089-191">Name</span></span>| <span data-ttu-id="3f089-192">类型</span><span class="sxs-lookup"><span data-stu-id="3f089-192">Type</span></span>| <span data-ttu-id="3f089-193">说明</span><span class="sxs-lookup"><span data-stu-id="3f089-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3f089-194">String</span><span class="sxs-lookup"><span data-stu-id="3f089-194">String</span></span>|<span data-ttu-id="3f089-195">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="3f089-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3f089-196">String</span><span class="sxs-lookup"><span data-stu-id="3f089-196">String</span></span>|<span data-ttu-id="3f089-197">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="3f089-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f089-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f089-198">Requirements</span></span>

|<span data-ttu-id="3f089-199">要求</span><span class="sxs-lookup"><span data-stu-id="3f089-199">Requirement</span></span>| <span data-ttu-id="3f089-200">值</span><span class="sxs-lookup"><span data-stu-id="3f089-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f089-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3f089-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f089-202">1.1</span><span class="sxs-lookup"><span data-stu-id="3f089-202">1.1</span></span>|
|[<span data-ttu-id="3f089-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3f089-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f089-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3f089-204">Compose or Read</span></span>|
