---
title: Office 命名空间-要求集1。2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0f955ed8279655b4ac92dc04871a1227b045f6ea
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165438"
---
# <a name="office"></a><span data-ttu-id="d5804-102">Office</span><span class="sxs-lookup"><span data-stu-id="d5804-102">Office</span></span>

<span data-ttu-id="d5804-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="d5804-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5804-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="d5804-105">Requirements</span></span>

|<span data-ttu-id="d5804-106">要求</span><span class="sxs-lookup"><span data-stu-id="d5804-106">Requirement</span></span>| <span data-ttu-id="d5804-107">值</span><span class="sxs-lookup"><span data-stu-id="d5804-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5804-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d5804-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d5804-109">1.1</span><span class="sxs-lookup"><span data-stu-id="d5804-109">1.1</span></span>|
|[<span data-ttu-id="d5804-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d5804-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d5804-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d5804-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d5804-112">属性</span><span class="sxs-lookup"><span data-stu-id="d5804-112">Properties</span></span>

| <span data-ttu-id="d5804-113">属性</span><span class="sxs-lookup"><span data-stu-id="d5804-113">Property</span></span> | <span data-ttu-id="d5804-114">型号</span><span class="sxs-lookup"><span data-stu-id="d5804-114">Modes</span></span> | <span data-ttu-id="d5804-115">返回类型</span><span class="sxs-lookup"><span data-stu-id="d5804-115">Return type</span></span> | <span data-ttu-id="d5804-116">最低</span><span class="sxs-lookup"><span data-stu-id="d5804-116">Minimum</span></span><br><span data-ttu-id="d5804-117">要求集</span><span class="sxs-lookup"><span data-stu-id="d5804-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d5804-118">context</span><span class="sxs-lookup"><span data-stu-id="d5804-118">context</span></span>](office.context.md) | <span data-ttu-id="d5804-119">撰写</span><span class="sxs-lookup"><span data-stu-id="d5804-119">Compose</span></span><br><span data-ttu-id="d5804-120">读取</span><span class="sxs-lookup"><span data-stu-id="d5804-120">Read</span></span> | [<span data-ttu-id="d5804-121">Context</span><span class="sxs-lookup"><span data-stu-id="d5804-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2) | [<span data-ttu-id="d5804-122">1.1</span><span class="sxs-lookup"><span data-stu-id="d5804-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="d5804-123">枚举</span><span class="sxs-lookup"><span data-stu-id="d5804-123">Enumerations</span></span>

| <span data-ttu-id="d5804-124">枚举</span><span class="sxs-lookup"><span data-stu-id="d5804-124">Enumeration</span></span> | <span data-ttu-id="d5804-125">型号</span><span class="sxs-lookup"><span data-stu-id="d5804-125">Modes</span></span> | <span data-ttu-id="d5804-126">返回类型</span><span class="sxs-lookup"><span data-stu-id="d5804-126">Return type</span></span> | <span data-ttu-id="d5804-127">最低</span><span class="sxs-lookup"><span data-stu-id="d5804-127">Minimum</span></span><br><span data-ttu-id="d5804-128">要求集</span><span class="sxs-lookup"><span data-stu-id="d5804-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d5804-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d5804-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d5804-130">撰写</span><span class="sxs-lookup"><span data-stu-id="d5804-130">Compose</span></span><br><span data-ttu-id="d5804-131">读取</span><span class="sxs-lookup"><span data-stu-id="d5804-131">Read</span></span> | <span data-ttu-id="d5804-132">String</span><span class="sxs-lookup"><span data-stu-id="d5804-132">String</span></span> | [<span data-ttu-id="d5804-133">1.1</span><span class="sxs-lookup"><span data-stu-id="d5804-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d5804-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d5804-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d5804-135">撰写</span><span class="sxs-lookup"><span data-stu-id="d5804-135">Compose</span></span><br><span data-ttu-id="d5804-136">读取</span><span class="sxs-lookup"><span data-stu-id="d5804-136">Read</span></span> | <span data-ttu-id="d5804-137">String</span><span class="sxs-lookup"><span data-stu-id="d5804-137">String</span></span> | [<span data-ttu-id="d5804-138">1.1</span><span class="sxs-lookup"><span data-stu-id="d5804-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d5804-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d5804-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d5804-140">撰写</span><span class="sxs-lookup"><span data-stu-id="d5804-140">Compose</span></span><br><span data-ttu-id="d5804-141">读取</span><span class="sxs-lookup"><span data-stu-id="d5804-141">Read</span></span> | <span data-ttu-id="d5804-142">String</span><span class="sxs-lookup"><span data-stu-id="d5804-142">String</span></span> | [<span data-ttu-id="d5804-143">1.1</span><span class="sxs-lookup"><span data-stu-id="d5804-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="d5804-144">命名空间</span><span class="sxs-lookup"><span data-stu-id="d5804-144">Namespaces</span></span>

<span data-ttu-id="d5804-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2)：包含许多特定于 Outlook 的`ItemType`枚举，例如`EntityType` `AttachmentType` `RecipientType` `ResponseType`、、、、、和`ItemNotificationMessageType`。</span><span class="sxs-lookup"><span data-stu-id="d5804-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="d5804-146">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="d5804-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d5804-147">AsyncResultStatus： String</span><span class="sxs-lookup"><span data-stu-id="d5804-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="d5804-148">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="d5804-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d5804-149">类型</span><span class="sxs-lookup"><span data-stu-id="d5804-149">Type</span></span>

*   <span data-ttu-id="d5804-150">String</span><span class="sxs-lookup"><span data-stu-id="d5804-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d5804-151">属性：</span><span class="sxs-lookup"><span data-stu-id="d5804-151">Properties:</span></span>

|<span data-ttu-id="d5804-152">名称</span><span class="sxs-lookup"><span data-stu-id="d5804-152">Name</span></span>| <span data-ttu-id="d5804-153">类型</span><span class="sxs-lookup"><span data-stu-id="d5804-153">Type</span></span>| <span data-ttu-id="d5804-154">说明</span><span class="sxs-lookup"><span data-stu-id="d5804-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d5804-155">String</span><span class="sxs-lookup"><span data-stu-id="d5804-155">String</span></span>|<span data-ttu-id="d5804-156">调用成功。</span><span class="sxs-lookup"><span data-stu-id="d5804-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d5804-157">String</span><span class="sxs-lookup"><span data-stu-id="d5804-157">String</span></span>|<span data-ttu-id="d5804-158">调用失败。</span><span class="sxs-lookup"><span data-stu-id="d5804-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5804-159">Requirements</span><span class="sxs-lookup"><span data-stu-id="d5804-159">Requirements</span></span>

|<span data-ttu-id="d5804-160">要求</span><span class="sxs-lookup"><span data-stu-id="d5804-160">Requirement</span></span>| <span data-ttu-id="d5804-161">值</span><span class="sxs-lookup"><span data-stu-id="d5804-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5804-162">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d5804-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d5804-163">1.1</span><span class="sxs-lookup"><span data-stu-id="d5804-163">1.1</span></span>|
|[<span data-ttu-id="d5804-164">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d5804-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d5804-165">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d5804-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d5804-166">CoercionType： String</span><span class="sxs-lookup"><span data-stu-id="d5804-166">CoercionType: String</span></span>

<span data-ttu-id="d5804-167">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="d5804-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d5804-168">类型</span><span class="sxs-lookup"><span data-stu-id="d5804-168">Type</span></span>

*   <span data-ttu-id="d5804-169">String</span><span class="sxs-lookup"><span data-stu-id="d5804-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d5804-170">属性：</span><span class="sxs-lookup"><span data-stu-id="d5804-170">Properties:</span></span>

|<span data-ttu-id="d5804-171">名称</span><span class="sxs-lookup"><span data-stu-id="d5804-171">Name</span></span>| <span data-ttu-id="d5804-172">类型</span><span class="sxs-lookup"><span data-stu-id="d5804-172">Type</span></span>| <span data-ttu-id="d5804-173">说明</span><span class="sxs-lookup"><span data-stu-id="d5804-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d5804-174">String</span><span class="sxs-lookup"><span data-stu-id="d5804-174">String</span></span>|<span data-ttu-id="d5804-175">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d5804-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d5804-176">String</span><span class="sxs-lookup"><span data-stu-id="d5804-176">String</span></span>|<span data-ttu-id="d5804-177">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="d5804-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5804-178">Requirements</span><span class="sxs-lookup"><span data-stu-id="d5804-178">Requirements</span></span>

|<span data-ttu-id="d5804-179">要求</span><span class="sxs-lookup"><span data-stu-id="d5804-179">Requirement</span></span>| <span data-ttu-id="d5804-180">值</span><span class="sxs-lookup"><span data-stu-id="d5804-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5804-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d5804-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d5804-182">1.1</span><span class="sxs-lookup"><span data-stu-id="d5804-182">1.1</span></span>|
|[<span data-ttu-id="d5804-183">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d5804-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d5804-184">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d5804-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d5804-185">SourceProperty： String</span><span class="sxs-lookup"><span data-stu-id="d5804-185">SourceProperty: String</span></span>

<span data-ttu-id="d5804-186">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="d5804-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d5804-187">类型</span><span class="sxs-lookup"><span data-stu-id="d5804-187">Type</span></span>

*   <span data-ttu-id="d5804-188">String</span><span class="sxs-lookup"><span data-stu-id="d5804-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d5804-189">属性：</span><span class="sxs-lookup"><span data-stu-id="d5804-189">Properties:</span></span>

|<span data-ttu-id="d5804-190">名称</span><span class="sxs-lookup"><span data-stu-id="d5804-190">Name</span></span>| <span data-ttu-id="d5804-191">类型</span><span class="sxs-lookup"><span data-stu-id="d5804-191">Type</span></span>| <span data-ttu-id="d5804-192">说明</span><span class="sxs-lookup"><span data-stu-id="d5804-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d5804-193">String</span><span class="sxs-lookup"><span data-stu-id="d5804-193">String</span></span>|<span data-ttu-id="d5804-194">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="d5804-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d5804-195">String</span><span class="sxs-lookup"><span data-stu-id="d5804-195">String</span></span>|<span data-ttu-id="d5804-196">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="d5804-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5804-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="d5804-197">Requirements</span></span>

|<span data-ttu-id="d5804-198">要求</span><span class="sxs-lookup"><span data-stu-id="d5804-198">Requirement</span></span>| <span data-ttu-id="d5804-199">值</span><span class="sxs-lookup"><span data-stu-id="d5804-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5804-200">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d5804-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d5804-201">1.1</span><span class="sxs-lookup"><span data-stu-id="d5804-201">1.1</span></span>|
|[<span data-ttu-id="d5804-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d5804-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d5804-203">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d5804-203">Compose or Read</span></span>|
