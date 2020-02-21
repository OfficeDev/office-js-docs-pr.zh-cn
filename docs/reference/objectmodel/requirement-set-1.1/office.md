---
title: Office 命名空间-要求集1。1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 68363f101b4c818853cc118e39d05784c56ef3ad
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165473"
---
# <a name="office"></a><span data-ttu-id="cebff-102">Office</span><span class="sxs-lookup"><span data-stu-id="cebff-102">Office</span></span>

<span data-ttu-id="cebff-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="cebff-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cebff-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="cebff-105">Requirements</span></span>

|<span data-ttu-id="cebff-106">要求</span><span class="sxs-lookup"><span data-stu-id="cebff-106">Requirement</span></span>| <span data-ttu-id="cebff-107">值</span><span class="sxs-lookup"><span data-stu-id="cebff-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cebff-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cebff-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cebff-109">1.1</span><span class="sxs-lookup"><span data-stu-id="cebff-109">1.1</span></span>|
|[<span data-ttu-id="cebff-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cebff-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cebff-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cebff-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cebff-112">属性</span><span class="sxs-lookup"><span data-stu-id="cebff-112">Properties</span></span>

| <span data-ttu-id="cebff-113">属性</span><span class="sxs-lookup"><span data-stu-id="cebff-113">Property</span></span> | <span data-ttu-id="cebff-114">型号</span><span class="sxs-lookup"><span data-stu-id="cebff-114">Modes</span></span> | <span data-ttu-id="cebff-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="cebff-115">Return type</span></span> | <span data-ttu-id="cebff-116">最低</span><span class="sxs-lookup"><span data-stu-id="cebff-116">Minimum</span></span><br><span data-ttu-id="cebff-117">要求集</span><span class="sxs-lookup"><span data-stu-id="cebff-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cebff-118">context</span><span class="sxs-lookup"><span data-stu-id="cebff-118">context</span></span>](office.context.md) | <span data-ttu-id="cebff-119">撰写</span><span class="sxs-lookup"><span data-stu-id="cebff-119">Compose</span></span><br><span data-ttu-id="cebff-120">读取</span><span class="sxs-lookup"><span data-stu-id="cebff-120">Read</span></span> | [<span data-ttu-id="cebff-121">Context</span><span class="sxs-lookup"><span data-stu-id="cebff-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="cebff-122">1.1</span><span class="sxs-lookup"><span data-stu-id="cebff-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="cebff-123">枚举</span><span class="sxs-lookup"><span data-stu-id="cebff-123">Enumerations</span></span>

| <span data-ttu-id="cebff-124">枚举</span><span class="sxs-lookup"><span data-stu-id="cebff-124">Enumeration</span></span> | <span data-ttu-id="cebff-125">型号</span><span class="sxs-lookup"><span data-stu-id="cebff-125">Modes</span></span> | <span data-ttu-id="cebff-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="cebff-126">Return type</span></span> | <span data-ttu-id="cebff-127">最低</span><span class="sxs-lookup"><span data-stu-id="cebff-127">Minimum</span></span><br><span data-ttu-id="cebff-128">要求集</span><span class="sxs-lookup"><span data-stu-id="cebff-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cebff-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="cebff-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="cebff-130">撰写</span><span class="sxs-lookup"><span data-stu-id="cebff-130">Compose</span></span><br><span data-ttu-id="cebff-131">读取</span><span class="sxs-lookup"><span data-stu-id="cebff-131">Read</span></span> | <span data-ttu-id="cebff-132">String</span><span class="sxs-lookup"><span data-stu-id="cebff-132">String</span></span> | [<span data-ttu-id="cebff-133">1.1</span><span class="sxs-lookup"><span data-stu-id="cebff-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cebff-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="cebff-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="cebff-135">撰写</span><span class="sxs-lookup"><span data-stu-id="cebff-135">Compose</span></span><br><span data-ttu-id="cebff-136">读取</span><span class="sxs-lookup"><span data-stu-id="cebff-136">Read</span></span> | <span data-ttu-id="cebff-137">String</span><span class="sxs-lookup"><span data-stu-id="cebff-137">String</span></span> | [<span data-ttu-id="cebff-138">1.1</span><span class="sxs-lookup"><span data-stu-id="cebff-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cebff-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="cebff-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="cebff-140">撰写</span><span class="sxs-lookup"><span data-stu-id="cebff-140">Compose</span></span><br><span data-ttu-id="cebff-141">读取</span><span class="sxs-lookup"><span data-stu-id="cebff-141">Read</span></span> | <span data-ttu-id="cebff-142">String</span><span class="sxs-lookup"><span data-stu-id="cebff-142">String</span></span> | [<span data-ttu-id="cebff-143">1.1</span><span class="sxs-lookup"><span data-stu-id="cebff-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="cebff-144">命名空间</span><span class="sxs-lookup"><span data-stu-id="cebff-144">Namespaces</span></span>

<span data-ttu-id="cebff-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="cebff-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="cebff-146">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="cebff-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="cebff-147">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="cebff-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="cebff-148">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="cebff-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="cebff-149">类型</span><span class="sxs-lookup"><span data-stu-id="cebff-149">Type</span></span>

*   <span data-ttu-id="cebff-150">String</span><span class="sxs-lookup"><span data-stu-id="cebff-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cebff-151">属性：</span><span class="sxs-lookup"><span data-stu-id="cebff-151">Properties:</span></span>

|<span data-ttu-id="cebff-152">名称</span><span class="sxs-lookup"><span data-stu-id="cebff-152">Name</span></span>| <span data-ttu-id="cebff-153">类型</span><span class="sxs-lookup"><span data-stu-id="cebff-153">Type</span></span>| <span data-ttu-id="cebff-154">说明</span><span class="sxs-lookup"><span data-stu-id="cebff-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="cebff-155">String</span><span class="sxs-lookup"><span data-stu-id="cebff-155">String</span></span>|<span data-ttu-id="cebff-156">调用成功。</span><span class="sxs-lookup"><span data-stu-id="cebff-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="cebff-157">String</span><span class="sxs-lookup"><span data-stu-id="cebff-157">String</span></span>|<span data-ttu-id="cebff-158">调用失败。</span><span class="sxs-lookup"><span data-stu-id="cebff-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cebff-159">Requirements</span><span class="sxs-lookup"><span data-stu-id="cebff-159">Requirements</span></span>

|<span data-ttu-id="cebff-160">要求</span><span class="sxs-lookup"><span data-stu-id="cebff-160">Requirement</span></span>| <span data-ttu-id="cebff-161">值</span><span class="sxs-lookup"><span data-stu-id="cebff-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="cebff-162">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cebff-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cebff-163">1.1</span><span class="sxs-lookup"><span data-stu-id="cebff-163">1.1</span></span>|
|[<span data-ttu-id="cebff-164">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cebff-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cebff-165">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cebff-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="cebff-166">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="cebff-166">CoercionType: String</span></span>

<span data-ttu-id="cebff-167">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="cebff-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cebff-168">类型</span><span class="sxs-lookup"><span data-stu-id="cebff-168">Type</span></span>

*   <span data-ttu-id="cebff-169">String</span><span class="sxs-lookup"><span data-stu-id="cebff-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cebff-170">属性：</span><span class="sxs-lookup"><span data-stu-id="cebff-170">Properties:</span></span>

|<span data-ttu-id="cebff-171">名称</span><span class="sxs-lookup"><span data-stu-id="cebff-171">Name</span></span>| <span data-ttu-id="cebff-172">类型</span><span class="sxs-lookup"><span data-stu-id="cebff-172">Type</span></span>| <span data-ttu-id="cebff-173">说明</span><span class="sxs-lookup"><span data-stu-id="cebff-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="cebff-174">String</span><span class="sxs-lookup"><span data-stu-id="cebff-174">String</span></span>|<span data-ttu-id="cebff-175">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="cebff-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="cebff-176">String</span><span class="sxs-lookup"><span data-stu-id="cebff-176">String</span></span>|<span data-ttu-id="cebff-177">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="cebff-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cebff-178">Requirements</span><span class="sxs-lookup"><span data-stu-id="cebff-178">Requirements</span></span>

|<span data-ttu-id="cebff-179">要求</span><span class="sxs-lookup"><span data-stu-id="cebff-179">Requirement</span></span>| <span data-ttu-id="cebff-180">值</span><span class="sxs-lookup"><span data-stu-id="cebff-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="cebff-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cebff-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cebff-182">1.1</span><span class="sxs-lookup"><span data-stu-id="cebff-182">1.1</span></span>|
|[<span data-ttu-id="cebff-183">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cebff-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cebff-184">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cebff-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="cebff-185">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="cebff-185">SourceProperty: String</span></span>

<span data-ttu-id="cebff-186">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="cebff-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cebff-187">类型</span><span class="sxs-lookup"><span data-stu-id="cebff-187">Type</span></span>

*   <span data-ttu-id="cebff-188">String</span><span class="sxs-lookup"><span data-stu-id="cebff-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cebff-189">属性：</span><span class="sxs-lookup"><span data-stu-id="cebff-189">Properties:</span></span>

|<span data-ttu-id="cebff-190">名称</span><span class="sxs-lookup"><span data-stu-id="cebff-190">Name</span></span>| <span data-ttu-id="cebff-191">类型</span><span class="sxs-lookup"><span data-stu-id="cebff-191">Type</span></span>| <span data-ttu-id="cebff-192">说明</span><span class="sxs-lookup"><span data-stu-id="cebff-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="cebff-193">String</span><span class="sxs-lookup"><span data-stu-id="cebff-193">String</span></span>|<span data-ttu-id="cebff-194">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="cebff-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="cebff-195">String</span><span class="sxs-lookup"><span data-stu-id="cebff-195">String</span></span>|<span data-ttu-id="cebff-196">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="cebff-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cebff-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="cebff-197">Requirements</span></span>

|<span data-ttu-id="cebff-198">要求</span><span class="sxs-lookup"><span data-stu-id="cebff-198">Requirement</span></span>| <span data-ttu-id="cebff-199">值</span><span class="sxs-lookup"><span data-stu-id="cebff-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="cebff-200">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cebff-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cebff-201">1.1</span><span class="sxs-lookup"><span data-stu-id="cebff-201">1.1</span></span>|
|[<span data-ttu-id="cebff-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cebff-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cebff-203">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cebff-203">Compose or Read</span></span>|