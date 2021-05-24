---
title: Office命名空间 - 要求集 1.3
description: Office邮箱 API 要求集 1.3 Outlook外接程序可用的命名空间成员。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f4aecf016e259141fd8adb2683864d4c36bdaf4b
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592001"
---
# <a name="office-mailbox-requirement-set-13"></a><span data-ttu-id="560e3-103">Office (邮箱要求集 1.3) </span><span class="sxs-lookup"><span data-stu-id="560e3-103">Office (Mailbox requirement set 1.3)</span></span>

<span data-ttu-id="560e3-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="560e3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="560e3-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="560e3-106">Requirements</span></span>

|<span data-ttu-id="560e3-107">要求</span><span class="sxs-lookup"><span data-stu-id="560e3-107">Requirement</span></span>| <span data-ttu-id="560e3-108">值</span><span class="sxs-lookup"><span data-stu-id="560e3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="560e3-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="560e3-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="560e3-110">1.1</span><span class="sxs-lookup"><span data-stu-id="560e3-110">1.1</span></span>|
|[<span data-ttu-id="560e3-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="560e3-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="560e3-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="560e3-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="560e3-113">属性</span><span class="sxs-lookup"><span data-stu-id="560e3-113">Properties</span></span>

| <span data-ttu-id="560e3-114">属性</span><span class="sxs-lookup"><span data-stu-id="560e3-114">Property</span></span> | <span data-ttu-id="560e3-115">模式</span><span class="sxs-lookup"><span data-stu-id="560e3-115">Modes</span></span> | <span data-ttu-id="560e3-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="560e3-116">Return type</span></span> | <span data-ttu-id="560e3-117">最小值</span><span class="sxs-lookup"><span data-stu-id="560e3-117">Minimum</span></span><br><span data-ttu-id="560e3-118">要求集</span><span class="sxs-lookup"><span data-stu-id="560e3-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="560e3-119">context</span><span class="sxs-lookup"><span data-stu-id="560e3-119">context</span></span>](office.context.md) | <span data-ttu-id="560e3-120">撰写</span><span class="sxs-lookup"><span data-stu-id="560e3-120">Compose</span></span><br><span data-ttu-id="560e3-121">阅读</span><span class="sxs-lookup"><span data-stu-id="560e3-121">Read</span></span> | [<span data-ttu-id="560e3-122">Context</span><span class="sxs-lookup"><span data-stu-id="560e3-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="560e3-123">1.1</span><span class="sxs-lookup"><span data-stu-id="560e3-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="560e3-124">枚举</span><span class="sxs-lookup"><span data-stu-id="560e3-124">Enumerations</span></span>

| <span data-ttu-id="560e3-125">枚举</span><span class="sxs-lookup"><span data-stu-id="560e3-125">Enumeration</span></span> | <span data-ttu-id="560e3-126">模式</span><span class="sxs-lookup"><span data-stu-id="560e3-126">Modes</span></span> | <span data-ttu-id="560e3-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="560e3-127">Return type</span></span> | <span data-ttu-id="560e3-128">最小值</span><span class="sxs-lookup"><span data-stu-id="560e3-128">Minimum</span></span><br><span data-ttu-id="560e3-129">要求集</span><span class="sxs-lookup"><span data-stu-id="560e3-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="560e3-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="560e3-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="560e3-131">撰写</span><span class="sxs-lookup"><span data-stu-id="560e3-131">Compose</span></span><br><span data-ttu-id="560e3-132">阅读</span><span class="sxs-lookup"><span data-stu-id="560e3-132">Read</span></span> | <span data-ttu-id="560e3-133">字符串</span><span class="sxs-lookup"><span data-stu-id="560e3-133">String</span></span> | [<span data-ttu-id="560e3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="560e3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="560e3-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="560e3-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="560e3-136">撰写</span><span class="sxs-lookup"><span data-stu-id="560e3-136">Compose</span></span><br><span data-ttu-id="560e3-137">阅读</span><span class="sxs-lookup"><span data-stu-id="560e3-137">Read</span></span> | <span data-ttu-id="560e3-138">字符串</span><span class="sxs-lookup"><span data-stu-id="560e3-138">String</span></span> | [<span data-ttu-id="560e3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="560e3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="560e3-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="560e3-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="560e3-141">撰写</span><span class="sxs-lookup"><span data-stu-id="560e3-141">Compose</span></span><br><span data-ttu-id="560e3-142">阅读</span><span class="sxs-lookup"><span data-stu-id="560e3-142">Read</span></span> | <span data-ttu-id="560e3-143">字符串</span><span class="sxs-lookup"><span data-stu-id="560e3-143">String</span></span> | [<span data-ttu-id="560e3-144">1.1</span><span class="sxs-lookup"><span data-stu-id="560e3-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="560e3-145">命名空间</span><span class="sxs-lookup"><span data-stu-id="560e3-145">Namespaces</span></span>

<span data-ttu-id="560e3-146">[MailboxEnums：](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3&preserve-view=true)包括许多Outlook枚举，例如、 `ItemType` `EntityType` 和 `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="560e3-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="560e3-147">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="560e3-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="560e3-148">AsyncResultStatus：String</span><span class="sxs-lookup"><span data-stu-id="560e3-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="560e3-149">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="560e3-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="560e3-150">类型</span><span class="sxs-lookup"><span data-stu-id="560e3-150">Type</span></span>

*   <span data-ttu-id="560e3-151">String</span><span class="sxs-lookup"><span data-stu-id="560e3-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="560e3-152">属性</span><span class="sxs-lookup"><span data-stu-id="560e3-152">Properties</span></span>

|<span data-ttu-id="560e3-153">名称</span><span class="sxs-lookup"><span data-stu-id="560e3-153">Name</span></span>| <span data-ttu-id="560e3-154">类型</span><span class="sxs-lookup"><span data-stu-id="560e3-154">Type</span></span>| <span data-ttu-id="560e3-155">描述</span><span class="sxs-lookup"><span data-stu-id="560e3-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="560e3-156">字符串</span><span class="sxs-lookup"><span data-stu-id="560e3-156">String</span></span>|<span data-ttu-id="560e3-157">调用成功。</span><span class="sxs-lookup"><span data-stu-id="560e3-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="560e3-158">字符串</span><span class="sxs-lookup"><span data-stu-id="560e3-158">String</span></span>|<span data-ttu-id="560e3-159">调用失败。</span><span class="sxs-lookup"><span data-stu-id="560e3-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="560e3-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="560e3-160">Requirements</span></span>

|<span data-ttu-id="560e3-161">要求</span><span class="sxs-lookup"><span data-stu-id="560e3-161">Requirement</span></span>| <span data-ttu-id="560e3-162">值</span><span class="sxs-lookup"><span data-stu-id="560e3-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="560e3-163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="560e3-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="560e3-164">1.1</span><span class="sxs-lookup"><span data-stu-id="560e3-164">1.1</span></span>|
|[<span data-ttu-id="560e3-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="560e3-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="560e3-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="560e3-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="560e3-167">CoercionType：String</span><span class="sxs-lookup"><span data-stu-id="560e3-167">CoercionType: String</span></span>

<span data-ttu-id="560e3-168">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="560e3-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="560e3-169">类型</span><span class="sxs-lookup"><span data-stu-id="560e3-169">Type</span></span>

*   <span data-ttu-id="560e3-170">String</span><span class="sxs-lookup"><span data-stu-id="560e3-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="560e3-171">属性</span><span class="sxs-lookup"><span data-stu-id="560e3-171">Properties</span></span>

|<span data-ttu-id="560e3-172">名称</span><span class="sxs-lookup"><span data-stu-id="560e3-172">Name</span></span>| <span data-ttu-id="560e3-173">类型</span><span class="sxs-lookup"><span data-stu-id="560e3-173">Type</span></span>| <span data-ttu-id="560e3-174">描述</span><span class="sxs-lookup"><span data-stu-id="560e3-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="560e3-175">字符串</span><span class="sxs-lookup"><span data-stu-id="560e3-175">String</span></span>|<span data-ttu-id="560e3-176">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="560e3-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="560e3-177">字符串</span><span class="sxs-lookup"><span data-stu-id="560e3-177">String</span></span>|<span data-ttu-id="560e3-178">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="560e3-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="560e3-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="560e3-179">Requirements</span></span>

|<span data-ttu-id="560e3-180">要求</span><span class="sxs-lookup"><span data-stu-id="560e3-180">Requirement</span></span>| <span data-ttu-id="560e3-181">值</span><span class="sxs-lookup"><span data-stu-id="560e3-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="560e3-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="560e3-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="560e3-183">1.1</span><span class="sxs-lookup"><span data-stu-id="560e3-183">1.1</span></span>|
|[<span data-ttu-id="560e3-184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="560e3-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="560e3-185">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="560e3-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="560e3-186">SourceProperty：String</span><span class="sxs-lookup"><span data-stu-id="560e3-186">SourceProperty: String</span></span>

<span data-ttu-id="560e3-187">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="560e3-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="560e3-188">类型</span><span class="sxs-lookup"><span data-stu-id="560e3-188">Type</span></span>

*   <span data-ttu-id="560e3-189">String</span><span class="sxs-lookup"><span data-stu-id="560e3-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="560e3-190">属性</span><span class="sxs-lookup"><span data-stu-id="560e3-190">Properties</span></span>

|<span data-ttu-id="560e3-191">名称</span><span class="sxs-lookup"><span data-stu-id="560e3-191">Name</span></span>| <span data-ttu-id="560e3-192">类型</span><span class="sxs-lookup"><span data-stu-id="560e3-192">Type</span></span>| <span data-ttu-id="560e3-193">描述</span><span class="sxs-lookup"><span data-stu-id="560e3-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="560e3-194">字符串</span><span class="sxs-lookup"><span data-stu-id="560e3-194">String</span></span>|<span data-ttu-id="560e3-195">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="560e3-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="560e3-196">String</span><span class="sxs-lookup"><span data-stu-id="560e3-196">String</span></span>|<span data-ttu-id="560e3-197">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="560e3-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="560e3-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="560e3-198">Requirements</span></span>

|<span data-ttu-id="560e3-199">要求</span><span class="sxs-lookup"><span data-stu-id="560e3-199">Requirement</span></span>| <span data-ttu-id="560e3-200">值</span><span class="sxs-lookup"><span data-stu-id="560e3-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="560e3-201">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="560e3-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="560e3-202">1.1</span><span class="sxs-lookup"><span data-stu-id="560e3-202">1.1</span></span>|
|[<span data-ttu-id="560e3-203">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="560e3-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="560e3-204">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="560e3-204">Compose or Read</span></span>|
