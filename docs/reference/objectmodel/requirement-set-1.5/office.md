---
title: Office命名空间 - 要求集 1.5
description: Office邮箱 API 要求集 1.5 Outlook外接程序可用的命名空间成员。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 46b70185ce983721c75093351e47a02eb8b9e7cd
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590853"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="41b67-103">Office (邮箱要求集 1.5) </span><span class="sxs-lookup"><span data-stu-id="41b67-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="41b67-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="41b67-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="41b67-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="41b67-106">Requirements</span></span>

|<span data-ttu-id="41b67-107">要求</span><span class="sxs-lookup"><span data-stu-id="41b67-107">Requirement</span></span>| <span data-ttu-id="41b67-108">值</span><span class="sxs-lookup"><span data-stu-id="41b67-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="41b67-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="41b67-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41b67-110">1.1</span><span class="sxs-lookup"><span data-stu-id="41b67-110">1.1</span></span>|
|[<span data-ttu-id="41b67-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="41b67-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41b67-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="41b67-113">属性</span><span class="sxs-lookup"><span data-stu-id="41b67-113">Properties</span></span>

| <span data-ttu-id="41b67-114">属性</span><span class="sxs-lookup"><span data-stu-id="41b67-114">Property</span></span> | <span data-ttu-id="41b67-115">模式</span><span class="sxs-lookup"><span data-stu-id="41b67-115">Modes</span></span> | <span data-ttu-id="41b67-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="41b67-116">Return type</span></span> | <span data-ttu-id="41b67-117">最小值</span><span class="sxs-lookup"><span data-stu-id="41b67-117">Minimum</span></span><br><span data-ttu-id="41b67-118">要求集</span><span class="sxs-lookup"><span data-stu-id="41b67-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="41b67-119">context</span><span class="sxs-lookup"><span data-stu-id="41b67-119">context</span></span>](office.context.md) | <span data-ttu-id="41b67-120">撰写</span><span class="sxs-lookup"><span data-stu-id="41b67-120">Compose</span></span><br><span data-ttu-id="41b67-121">阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-121">Read</span></span> | [<span data-ttu-id="41b67-122">Context</span><span class="sxs-lookup"><span data-stu-id="41b67-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="41b67-123">1.1</span><span class="sxs-lookup"><span data-stu-id="41b67-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="41b67-124">枚举</span><span class="sxs-lookup"><span data-stu-id="41b67-124">Enumerations</span></span>

| <span data-ttu-id="41b67-125">枚举</span><span class="sxs-lookup"><span data-stu-id="41b67-125">Enumeration</span></span> | <span data-ttu-id="41b67-126">模式</span><span class="sxs-lookup"><span data-stu-id="41b67-126">Modes</span></span> | <span data-ttu-id="41b67-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="41b67-127">Return type</span></span> | <span data-ttu-id="41b67-128">最小值</span><span class="sxs-lookup"><span data-stu-id="41b67-128">Minimum</span></span><br><span data-ttu-id="41b67-129">要求集</span><span class="sxs-lookup"><span data-stu-id="41b67-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="41b67-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="41b67-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="41b67-131">撰写</span><span class="sxs-lookup"><span data-stu-id="41b67-131">Compose</span></span><br><span data-ttu-id="41b67-132">阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-132">Read</span></span> | <span data-ttu-id="41b67-133">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-133">String</span></span> | [<span data-ttu-id="41b67-134">1.1</span><span class="sxs-lookup"><span data-stu-id="41b67-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="41b67-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="41b67-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="41b67-136">撰写</span><span class="sxs-lookup"><span data-stu-id="41b67-136">Compose</span></span><br><span data-ttu-id="41b67-137">阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-137">Read</span></span> | <span data-ttu-id="41b67-138">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-138">String</span></span> | [<span data-ttu-id="41b67-139">1.1</span><span class="sxs-lookup"><span data-stu-id="41b67-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="41b67-140">EventType</span><span class="sxs-lookup"><span data-stu-id="41b67-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="41b67-141">撰写</span><span class="sxs-lookup"><span data-stu-id="41b67-141">Compose</span></span><br><span data-ttu-id="41b67-142">阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-142">Read</span></span> | <span data-ttu-id="41b67-143">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-143">String</span></span> | [<span data-ttu-id="41b67-144">1.5</span><span class="sxs-lookup"><span data-stu-id="41b67-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="41b67-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="41b67-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="41b67-146">撰写</span><span class="sxs-lookup"><span data-stu-id="41b67-146">Compose</span></span><br><span data-ttu-id="41b67-147">阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-147">Read</span></span> | <span data-ttu-id="41b67-148">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-148">String</span></span> | [<span data-ttu-id="41b67-149">1.1</span><span class="sxs-lookup"><span data-stu-id="41b67-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="41b67-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="41b67-150">Namespaces</span></span>

<span data-ttu-id="41b67-151">[MailboxEnums：](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true)包括许多Outlook枚举，例如、 `ItemType` `EntityType` 和 `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="41b67-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="41b67-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="41b67-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="41b67-153">AsyncResultStatus：String</span><span class="sxs-lookup"><span data-stu-id="41b67-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="41b67-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="41b67-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="41b67-155">类型</span><span class="sxs-lookup"><span data-stu-id="41b67-155">Type</span></span>

*   <span data-ttu-id="41b67-156">String</span><span class="sxs-lookup"><span data-stu-id="41b67-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41b67-157">属性</span><span class="sxs-lookup"><span data-stu-id="41b67-157">Properties</span></span>

|<span data-ttu-id="41b67-158">名称</span><span class="sxs-lookup"><span data-stu-id="41b67-158">Name</span></span>| <span data-ttu-id="41b67-159">类型</span><span class="sxs-lookup"><span data-stu-id="41b67-159">Type</span></span>| <span data-ttu-id="41b67-160">描述</span><span class="sxs-lookup"><span data-stu-id="41b67-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="41b67-161">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-161">String</span></span>|<span data-ttu-id="41b67-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="41b67-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="41b67-163">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-163">String</span></span>|<span data-ttu-id="41b67-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="41b67-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41b67-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="41b67-165">Requirements</span></span>

|<span data-ttu-id="41b67-166">要求</span><span class="sxs-lookup"><span data-stu-id="41b67-166">Requirement</span></span>| <span data-ttu-id="41b67-167">值</span><span class="sxs-lookup"><span data-stu-id="41b67-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="41b67-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="41b67-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41b67-169">1.1</span><span class="sxs-lookup"><span data-stu-id="41b67-169">1.1</span></span>|
|[<span data-ttu-id="41b67-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="41b67-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41b67-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="41b67-172">CoercionType：String</span><span class="sxs-lookup"><span data-stu-id="41b67-172">CoercionType: String</span></span>

<span data-ttu-id="41b67-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="41b67-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="41b67-174">类型</span><span class="sxs-lookup"><span data-stu-id="41b67-174">Type</span></span>

*   <span data-ttu-id="41b67-175">String</span><span class="sxs-lookup"><span data-stu-id="41b67-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41b67-176">属性</span><span class="sxs-lookup"><span data-stu-id="41b67-176">Properties</span></span>

|<span data-ttu-id="41b67-177">名称</span><span class="sxs-lookup"><span data-stu-id="41b67-177">Name</span></span>| <span data-ttu-id="41b67-178">类型</span><span class="sxs-lookup"><span data-stu-id="41b67-178">Type</span></span>| <span data-ttu-id="41b67-179">描述</span><span class="sxs-lookup"><span data-stu-id="41b67-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="41b67-180">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-180">String</span></span>|<span data-ttu-id="41b67-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="41b67-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="41b67-182">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-182">String</span></span>|<span data-ttu-id="41b67-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="41b67-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41b67-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="41b67-184">Requirements</span></span>

|<span data-ttu-id="41b67-185">要求</span><span class="sxs-lookup"><span data-stu-id="41b67-185">Requirement</span></span>| <span data-ttu-id="41b67-186">值</span><span class="sxs-lookup"><span data-stu-id="41b67-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="41b67-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="41b67-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41b67-188">1.1</span><span class="sxs-lookup"><span data-stu-id="41b67-188">1.1</span></span>|
|[<span data-ttu-id="41b67-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="41b67-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41b67-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="41b67-191">EventType：String</span><span class="sxs-lookup"><span data-stu-id="41b67-191">EventType: String</span></span>

<span data-ttu-id="41b67-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="41b67-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="41b67-193">类型</span><span class="sxs-lookup"><span data-stu-id="41b67-193">Type</span></span>

*   <span data-ttu-id="41b67-194">String</span><span class="sxs-lookup"><span data-stu-id="41b67-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41b67-195">属性</span><span class="sxs-lookup"><span data-stu-id="41b67-195">Properties</span></span>

| <span data-ttu-id="41b67-196">名称</span><span class="sxs-lookup"><span data-stu-id="41b67-196">Name</span></span> | <span data-ttu-id="41b67-197">类型</span><span class="sxs-lookup"><span data-stu-id="41b67-197">Type</span></span> | <span data-ttu-id="41b67-198">描述</span><span class="sxs-lookup"><span data-stu-id="41b67-198">Description</span></span> | <span data-ttu-id="41b67-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="41b67-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="41b67-200">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-200">String</span></span> | <span data-ttu-id="41b67-201">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="41b67-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="41b67-202">1.5</span><span class="sxs-lookup"><span data-stu-id="41b67-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="41b67-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="41b67-203">Requirements</span></span>

|<span data-ttu-id="41b67-204">要求</span><span class="sxs-lookup"><span data-stu-id="41b67-204">Requirement</span></span>| <span data-ttu-id="41b67-205">值</span><span class="sxs-lookup"><span data-stu-id="41b67-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="41b67-206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="41b67-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41b67-207">1.5</span><span class="sxs-lookup"><span data-stu-id="41b67-207">1.5</span></span> |
|[<span data-ttu-id="41b67-208">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="41b67-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41b67-209">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="41b67-210">SourceProperty：String</span><span class="sxs-lookup"><span data-stu-id="41b67-210">SourceProperty: String</span></span>

<span data-ttu-id="41b67-211">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="41b67-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="41b67-212">类型</span><span class="sxs-lookup"><span data-stu-id="41b67-212">Type</span></span>

*   <span data-ttu-id="41b67-213">String</span><span class="sxs-lookup"><span data-stu-id="41b67-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41b67-214">属性</span><span class="sxs-lookup"><span data-stu-id="41b67-214">Properties</span></span>

|<span data-ttu-id="41b67-215">名称</span><span class="sxs-lookup"><span data-stu-id="41b67-215">Name</span></span>| <span data-ttu-id="41b67-216">类型</span><span class="sxs-lookup"><span data-stu-id="41b67-216">Type</span></span>| <span data-ttu-id="41b67-217">描述</span><span class="sxs-lookup"><span data-stu-id="41b67-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="41b67-218">字符串</span><span class="sxs-lookup"><span data-stu-id="41b67-218">String</span></span>|<span data-ttu-id="41b67-219">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="41b67-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="41b67-220">String</span><span class="sxs-lookup"><span data-stu-id="41b67-220">String</span></span>|<span data-ttu-id="41b67-221">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="41b67-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41b67-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="41b67-222">Requirements</span></span>

|<span data-ttu-id="41b67-223">要求</span><span class="sxs-lookup"><span data-stu-id="41b67-223">Requirement</span></span>| <span data-ttu-id="41b67-224">值</span><span class="sxs-lookup"><span data-stu-id="41b67-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="41b67-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="41b67-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41b67-226">1.1</span><span class="sxs-lookup"><span data-stu-id="41b67-226">1.1</span></span>|
|[<span data-ttu-id="41b67-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="41b67-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41b67-228">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="41b67-228">Compose or Read</span></span>|
