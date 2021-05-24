---
title: Office命名空间 - 要求集 1.6
description: Office邮箱 API 要求集 1.6 Outlook外接程序可用的命名空间成员。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 40cdb7de0678007b93b9251e7f1e2921ed857338
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590832"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="243c0-103">Office (邮箱要求集 1.6) </span><span class="sxs-lookup"><span data-stu-id="243c0-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="243c0-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="243c0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="243c0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="243c0-106">Requirements</span></span>

|<span data-ttu-id="243c0-107">要求</span><span class="sxs-lookup"><span data-stu-id="243c0-107">Requirement</span></span>| <span data-ttu-id="243c0-108">值</span><span class="sxs-lookup"><span data-stu-id="243c0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="243c0-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="243c0-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="243c0-110">1.1</span><span class="sxs-lookup"><span data-stu-id="243c0-110">1.1</span></span>|
|[<span data-ttu-id="243c0-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="243c0-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="243c0-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="243c0-113">属性</span><span class="sxs-lookup"><span data-stu-id="243c0-113">Properties</span></span>

| <span data-ttu-id="243c0-114">属性</span><span class="sxs-lookup"><span data-stu-id="243c0-114">Property</span></span> | <span data-ttu-id="243c0-115">模式</span><span class="sxs-lookup"><span data-stu-id="243c0-115">Modes</span></span> | <span data-ttu-id="243c0-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="243c0-116">Return type</span></span> | <span data-ttu-id="243c0-117">最小值</span><span class="sxs-lookup"><span data-stu-id="243c0-117">Minimum</span></span><br><span data-ttu-id="243c0-118">要求集</span><span class="sxs-lookup"><span data-stu-id="243c0-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="243c0-119">context</span><span class="sxs-lookup"><span data-stu-id="243c0-119">context</span></span>](office.context.md) | <span data-ttu-id="243c0-120">撰写</span><span class="sxs-lookup"><span data-stu-id="243c0-120">Compose</span></span><br><span data-ttu-id="243c0-121">阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-121">Read</span></span> | [<span data-ttu-id="243c0-122">Context</span><span class="sxs-lookup"><span data-stu-id="243c0-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="243c0-123">1.1</span><span class="sxs-lookup"><span data-stu-id="243c0-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="243c0-124">枚举</span><span class="sxs-lookup"><span data-stu-id="243c0-124">Enumerations</span></span>

| <span data-ttu-id="243c0-125">枚举</span><span class="sxs-lookup"><span data-stu-id="243c0-125">Enumeration</span></span> | <span data-ttu-id="243c0-126">模式</span><span class="sxs-lookup"><span data-stu-id="243c0-126">Modes</span></span> | <span data-ttu-id="243c0-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="243c0-127">Return type</span></span> | <span data-ttu-id="243c0-128">最小值</span><span class="sxs-lookup"><span data-stu-id="243c0-128">Minimum</span></span><br><span data-ttu-id="243c0-129">要求集</span><span class="sxs-lookup"><span data-stu-id="243c0-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="243c0-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="243c0-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="243c0-131">撰写</span><span class="sxs-lookup"><span data-stu-id="243c0-131">Compose</span></span><br><span data-ttu-id="243c0-132">阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-132">Read</span></span> | <span data-ttu-id="243c0-133">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-133">String</span></span> | [<span data-ttu-id="243c0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="243c0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="243c0-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="243c0-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="243c0-136">撰写</span><span class="sxs-lookup"><span data-stu-id="243c0-136">Compose</span></span><br><span data-ttu-id="243c0-137">阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-137">Read</span></span> | <span data-ttu-id="243c0-138">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-138">String</span></span> | [<span data-ttu-id="243c0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="243c0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="243c0-140">EventType</span><span class="sxs-lookup"><span data-stu-id="243c0-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="243c0-141">撰写</span><span class="sxs-lookup"><span data-stu-id="243c0-141">Compose</span></span><br><span data-ttu-id="243c0-142">阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-142">Read</span></span> | <span data-ttu-id="243c0-143">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-143">String</span></span> | [<span data-ttu-id="243c0-144">1.5</span><span class="sxs-lookup"><span data-stu-id="243c0-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="243c0-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="243c0-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="243c0-146">撰写</span><span class="sxs-lookup"><span data-stu-id="243c0-146">Compose</span></span><br><span data-ttu-id="243c0-147">阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-147">Read</span></span> | <span data-ttu-id="243c0-148">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-148">String</span></span> | [<span data-ttu-id="243c0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="243c0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="243c0-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="243c0-150">Namespaces</span></span>

<span data-ttu-id="243c0-151">[MailboxEnums：](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true)包括许多Outlook枚举，例如、 `ItemType` `EntityType` 和 `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="243c0-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="243c0-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="243c0-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="243c0-153">AsyncResultStatus：String</span><span class="sxs-lookup"><span data-stu-id="243c0-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="243c0-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="243c0-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="243c0-155">类型</span><span class="sxs-lookup"><span data-stu-id="243c0-155">Type</span></span>

*   <span data-ttu-id="243c0-156">String</span><span class="sxs-lookup"><span data-stu-id="243c0-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243c0-157">属性</span><span class="sxs-lookup"><span data-stu-id="243c0-157">Properties</span></span>

|<span data-ttu-id="243c0-158">名称</span><span class="sxs-lookup"><span data-stu-id="243c0-158">Name</span></span>| <span data-ttu-id="243c0-159">类型</span><span class="sxs-lookup"><span data-stu-id="243c0-159">Type</span></span>| <span data-ttu-id="243c0-160">描述</span><span class="sxs-lookup"><span data-stu-id="243c0-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="243c0-161">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-161">String</span></span>|<span data-ttu-id="243c0-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="243c0-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="243c0-163">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-163">String</span></span>|<span data-ttu-id="243c0-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="243c0-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243c0-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="243c0-165">Requirements</span></span>

|<span data-ttu-id="243c0-166">要求</span><span class="sxs-lookup"><span data-stu-id="243c0-166">Requirement</span></span>| <span data-ttu-id="243c0-167">值</span><span class="sxs-lookup"><span data-stu-id="243c0-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="243c0-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="243c0-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="243c0-169">1.1</span><span class="sxs-lookup"><span data-stu-id="243c0-169">1.1</span></span>|
|[<span data-ttu-id="243c0-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="243c0-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="243c0-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="243c0-172">CoercionType：String</span><span class="sxs-lookup"><span data-stu-id="243c0-172">CoercionType: String</span></span>

<span data-ttu-id="243c0-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="243c0-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="243c0-174">类型</span><span class="sxs-lookup"><span data-stu-id="243c0-174">Type</span></span>

*   <span data-ttu-id="243c0-175">String</span><span class="sxs-lookup"><span data-stu-id="243c0-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243c0-176">属性</span><span class="sxs-lookup"><span data-stu-id="243c0-176">Properties</span></span>

|<span data-ttu-id="243c0-177">名称</span><span class="sxs-lookup"><span data-stu-id="243c0-177">Name</span></span>| <span data-ttu-id="243c0-178">类型</span><span class="sxs-lookup"><span data-stu-id="243c0-178">Type</span></span>| <span data-ttu-id="243c0-179">描述</span><span class="sxs-lookup"><span data-stu-id="243c0-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="243c0-180">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-180">String</span></span>|<span data-ttu-id="243c0-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="243c0-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="243c0-182">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-182">String</span></span>|<span data-ttu-id="243c0-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="243c0-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243c0-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="243c0-184">Requirements</span></span>

|<span data-ttu-id="243c0-185">要求</span><span class="sxs-lookup"><span data-stu-id="243c0-185">Requirement</span></span>| <span data-ttu-id="243c0-186">值</span><span class="sxs-lookup"><span data-stu-id="243c0-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="243c0-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="243c0-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="243c0-188">1.1</span><span class="sxs-lookup"><span data-stu-id="243c0-188">1.1</span></span>|
|[<span data-ttu-id="243c0-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="243c0-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="243c0-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="243c0-191">EventType：String</span><span class="sxs-lookup"><span data-stu-id="243c0-191">EventType: String</span></span>

<span data-ttu-id="243c0-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="243c0-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="243c0-193">类型</span><span class="sxs-lookup"><span data-stu-id="243c0-193">Type</span></span>

*   <span data-ttu-id="243c0-194">String</span><span class="sxs-lookup"><span data-stu-id="243c0-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243c0-195">属性</span><span class="sxs-lookup"><span data-stu-id="243c0-195">Properties</span></span>

| <span data-ttu-id="243c0-196">名称</span><span class="sxs-lookup"><span data-stu-id="243c0-196">Name</span></span> | <span data-ttu-id="243c0-197">类型</span><span class="sxs-lookup"><span data-stu-id="243c0-197">Type</span></span> | <span data-ttu-id="243c0-198">描述</span><span class="sxs-lookup"><span data-stu-id="243c0-198">Description</span></span> | <span data-ttu-id="243c0-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="243c0-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="243c0-200">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-200">String</span></span> | <span data-ttu-id="243c0-201">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="243c0-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="243c0-202">1.5</span><span class="sxs-lookup"><span data-stu-id="243c0-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="243c0-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="243c0-203">Requirements</span></span>

|<span data-ttu-id="243c0-204">要求</span><span class="sxs-lookup"><span data-stu-id="243c0-204">Requirement</span></span>| <span data-ttu-id="243c0-205">值</span><span class="sxs-lookup"><span data-stu-id="243c0-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="243c0-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="243c0-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="243c0-207">1.5</span><span class="sxs-lookup"><span data-stu-id="243c0-207">1.5</span></span> |
|[<span data-ttu-id="243c0-208">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="243c0-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="243c0-209">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="243c0-210">SourceProperty：String</span><span class="sxs-lookup"><span data-stu-id="243c0-210">SourceProperty: String</span></span>

<span data-ttu-id="243c0-211">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="243c0-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="243c0-212">类型</span><span class="sxs-lookup"><span data-stu-id="243c0-212">Type</span></span>

*   <span data-ttu-id="243c0-213">String</span><span class="sxs-lookup"><span data-stu-id="243c0-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="243c0-214">属性</span><span class="sxs-lookup"><span data-stu-id="243c0-214">Properties</span></span>

|<span data-ttu-id="243c0-215">名称</span><span class="sxs-lookup"><span data-stu-id="243c0-215">Name</span></span>| <span data-ttu-id="243c0-216">类型</span><span class="sxs-lookup"><span data-stu-id="243c0-216">Type</span></span>| <span data-ttu-id="243c0-217">描述</span><span class="sxs-lookup"><span data-stu-id="243c0-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="243c0-218">字符串</span><span class="sxs-lookup"><span data-stu-id="243c0-218">String</span></span>|<span data-ttu-id="243c0-219">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="243c0-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="243c0-220">String</span><span class="sxs-lookup"><span data-stu-id="243c0-220">String</span></span>|<span data-ttu-id="243c0-221">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="243c0-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="243c0-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="243c0-222">Requirements</span></span>

|<span data-ttu-id="243c0-223">要求</span><span class="sxs-lookup"><span data-stu-id="243c0-223">Requirement</span></span>| <span data-ttu-id="243c0-224">值</span><span class="sxs-lookup"><span data-stu-id="243c0-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="243c0-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="243c0-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="243c0-226">1.1</span><span class="sxs-lookup"><span data-stu-id="243c0-226">1.1</span></span>|
|[<span data-ttu-id="243c0-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="243c0-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="243c0-228">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="243c0-228">Compose or Read</span></span>|
