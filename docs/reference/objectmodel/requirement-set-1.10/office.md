---
title: Office命名空间 - 要求集 1.10
description: Office邮箱 API 要求集 1.10 Outlook外接程序可用的命名空间成员。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e7b7ab9127ebf8ce9b7394d348144fe63b47de6c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592029"
---
# <a name="office-mailbox-requirement-set-110"></a><span data-ttu-id="af90c-103">Office (邮箱要求集 1.10) </span><span class="sxs-lookup"><span data-stu-id="af90c-103">Office (Mailbox requirement set 1.10)</span></span>

<span data-ttu-id="af90c-p101">该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="af90c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="af90c-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="af90c-106">Requirements</span></span>

|<span data-ttu-id="af90c-107">要求</span><span class="sxs-lookup"><span data-stu-id="af90c-107">Requirement</span></span>| <span data-ttu-id="af90c-108">值</span><span class="sxs-lookup"><span data-stu-id="af90c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="af90c-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="af90c-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af90c-110">1.1</span><span class="sxs-lookup"><span data-stu-id="af90c-110">1.1</span></span>|
|[<span data-ttu-id="af90c-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="af90c-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af90c-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="af90c-113">属性</span><span class="sxs-lookup"><span data-stu-id="af90c-113">Properties</span></span>

| <span data-ttu-id="af90c-114">属性</span><span class="sxs-lookup"><span data-stu-id="af90c-114">Property</span></span> | <span data-ttu-id="af90c-115">模式</span><span class="sxs-lookup"><span data-stu-id="af90c-115">Modes</span></span> | <span data-ttu-id="af90c-116">返回类型</span><span class="sxs-lookup"><span data-stu-id="af90c-116">Return type</span></span> | <span data-ttu-id="af90c-117">最小值</span><span class="sxs-lookup"><span data-stu-id="af90c-117">Minimum</span></span><br><span data-ttu-id="af90c-118">要求集</span><span class="sxs-lookup"><span data-stu-id="af90c-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="af90c-119">context</span><span class="sxs-lookup"><span data-stu-id="af90c-119">context</span></span>](office.context.md) | <span data-ttu-id="af90c-120">撰写</span><span class="sxs-lookup"><span data-stu-id="af90c-120">Compose</span></span><br><span data-ttu-id="af90c-121">阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-121">Read</span></span> | [<span data-ttu-id="af90c-122">Context</span><span class="sxs-lookup"><span data-stu-id="af90c-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="af90c-123">1.1</span><span class="sxs-lookup"><span data-stu-id="af90c-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="af90c-124">枚举</span><span class="sxs-lookup"><span data-stu-id="af90c-124">Enumerations</span></span>

| <span data-ttu-id="af90c-125">枚举</span><span class="sxs-lookup"><span data-stu-id="af90c-125">Enumeration</span></span> | <span data-ttu-id="af90c-126">模式</span><span class="sxs-lookup"><span data-stu-id="af90c-126">Modes</span></span> | <span data-ttu-id="af90c-127">返回类型</span><span class="sxs-lookup"><span data-stu-id="af90c-127">Return type</span></span> | <span data-ttu-id="af90c-128">最小值</span><span class="sxs-lookup"><span data-stu-id="af90c-128">Minimum</span></span><br><span data-ttu-id="af90c-129">要求集</span><span class="sxs-lookup"><span data-stu-id="af90c-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="af90c-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="af90c-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="af90c-131">撰写</span><span class="sxs-lookup"><span data-stu-id="af90c-131">Compose</span></span><br><span data-ttu-id="af90c-132">阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-132">Read</span></span> | <span data-ttu-id="af90c-133">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-133">String</span></span> | [<span data-ttu-id="af90c-134">1.1</span><span class="sxs-lookup"><span data-stu-id="af90c-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af90c-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="af90c-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="af90c-136">撰写</span><span class="sxs-lookup"><span data-stu-id="af90c-136">Compose</span></span><br><span data-ttu-id="af90c-137">阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-137">Read</span></span> | <span data-ttu-id="af90c-138">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-138">String</span></span> | [<span data-ttu-id="af90c-139">1.1</span><span class="sxs-lookup"><span data-stu-id="af90c-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af90c-140">EventType</span><span class="sxs-lookup"><span data-stu-id="af90c-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="af90c-141">撰写</span><span class="sxs-lookup"><span data-stu-id="af90c-141">Compose</span></span><br><span data-ttu-id="af90c-142">阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-142">Read</span></span> | <span data-ttu-id="af90c-143">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-143">String</span></span> | [<span data-ttu-id="af90c-144">1.5</span><span class="sxs-lookup"><span data-stu-id="af90c-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="af90c-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="af90c-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="af90c-146">撰写</span><span class="sxs-lookup"><span data-stu-id="af90c-146">Compose</span></span><br><span data-ttu-id="af90c-147">阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-147">Read</span></span> | <span data-ttu-id="af90c-148">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-148">String</span></span> | [<span data-ttu-id="af90c-149">1.1</span><span class="sxs-lookup"><span data-stu-id="af90c-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="af90c-150">命名空间</span><span class="sxs-lookup"><span data-stu-id="af90c-150">Namespaces</span></span>

<span data-ttu-id="af90c-151">[MailboxEnums：](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true)包括许多Outlook枚举，例如、 `ItemType` `EntityType` 和 `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="af90c-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="af90c-152">枚举详细信息</span><span class="sxs-lookup"><span data-stu-id="af90c-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="af90c-153">AsyncResultStatus：String</span><span class="sxs-lookup"><span data-stu-id="af90c-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="af90c-154">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="af90c-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="af90c-155">类型</span><span class="sxs-lookup"><span data-stu-id="af90c-155">Type</span></span>

*   <span data-ttu-id="af90c-156">String</span><span class="sxs-lookup"><span data-stu-id="af90c-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="af90c-157">属性</span><span class="sxs-lookup"><span data-stu-id="af90c-157">Properties</span></span>

|<span data-ttu-id="af90c-158">名称</span><span class="sxs-lookup"><span data-stu-id="af90c-158">Name</span></span>| <span data-ttu-id="af90c-159">类型</span><span class="sxs-lookup"><span data-stu-id="af90c-159">Type</span></span>| <span data-ttu-id="af90c-160">描述</span><span class="sxs-lookup"><span data-stu-id="af90c-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="af90c-161">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-161">String</span></span>|<span data-ttu-id="af90c-162">调用成功。</span><span class="sxs-lookup"><span data-stu-id="af90c-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="af90c-163">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-163">String</span></span>|<span data-ttu-id="af90c-164">调用失败。</span><span class="sxs-lookup"><span data-stu-id="af90c-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af90c-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="af90c-165">Requirements</span></span>

|<span data-ttu-id="af90c-166">要求</span><span class="sxs-lookup"><span data-stu-id="af90c-166">Requirement</span></span>| <span data-ttu-id="af90c-167">值</span><span class="sxs-lookup"><span data-stu-id="af90c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="af90c-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="af90c-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af90c-169">1.1</span><span class="sxs-lookup"><span data-stu-id="af90c-169">1.1</span></span>|
|[<span data-ttu-id="af90c-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="af90c-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af90c-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="af90c-172">CoercionType：String</span><span class="sxs-lookup"><span data-stu-id="af90c-172">CoercionType: String</span></span>

<span data-ttu-id="af90c-173">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="af90c-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="af90c-174">类型</span><span class="sxs-lookup"><span data-stu-id="af90c-174">Type</span></span>

*   <span data-ttu-id="af90c-175">String</span><span class="sxs-lookup"><span data-stu-id="af90c-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="af90c-176">属性</span><span class="sxs-lookup"><span data-stu-id="af90c-176">Properties</span></span>

|<span data-ttu-id="af90c-177">名称</span><span class="sxs-lookup"><span data-stu-id="af90c-177">Name</span></span>| <span data-ttu-id="af90c-178">类型</span><span class="sxs-lookup"><span data-stu-id="af90c-178">Type</span></span>| <span data-ttu-id="af90c-179">描述</span><span class="sxs-lookup"><span data-stu-id="af90c-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="af90c-180">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-180">String</span></span>|<span data-ttu-id="af90c-181">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="af90c-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="af90c-182">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-182">String</span></span>|<span data-ttu-id="af90c-183">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="af90c-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af90c-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="af90c-184">Requirements</span></span>

|<span data-ttu-id="af90c-185">要求</span><span class="sxs-lookup"><span data-stu-id="af90c-185">Requirement</span></span>| <span data-ttu-id="af90c-186">值</span><span class="sxs-lookup"><span data-stu-id="af90c-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="af90c-187">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="af90c-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af90c-188">1.1</span><span class="sxs-lookup"><span data-stu-id="af90c-188">1.1</span></span>|
|[<span data-ttu-id="af90c-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="af90c-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af90c-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="af90c-191">EventType：String</span><span class="sxs-lookup"><span data-stu-id="af90c-191">EventType: String</span></span>

<span data-ttu-id="af90c-192">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="af90c-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="af90c-193">类型</span><span class="sxs-lookup"><span data-stu-id="af90c-193">Type</span></span>

*   <span data-ttu-id="af90c-194">String</span><span class="sxs-lookup"><span data-stu-id="af90c-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="af90c-195">属性</span><span class="sxs-lookup"><span data-stu-id="af90c-195">Properties</span></span>

| <span data-ttu-id="af90c-196">名称</span><span class="sxs-lookup"><span data-stu-id="af90c-196">Name</span></span> | <span data-ttu-id="af90c-197">类型</span><span class="sxs-lookup"><span data-stu-id="af90c-197">Type</span></span> | <span data-ttu-id="af90c-198">描述</span><span class="sxs-lookup"><span data-stu-id="af90c-198">Description</span></span> | <span data-ttu-id="af90c-199">最低要求集</span><span class="sxs-lookup"><span data-stu-id="af90c-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="af90c-200">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-200">String</span></span> | <span data-ttu-id="af90c-201">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="af90c-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="af90c-202">1.7</span><span class="sxs-lookup"><span data-stu-id="af90c-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="af90c-203">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-203">String</span></span> | <span data-ttu-id="af90c-204">已将附件添加到项目或已从项目删除附件。</span><span class="sxs-lookup"><span data-stu-id="af90c-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="af90c-205">1.8</span><span class="sxs-lookup"><span data-stu-id="af90c-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="af90c-206">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-206">String</span></span> | <span data-ttu-id="af90c-207">所选约会的位置已更改。</span><span class="sxs-lookup"><span data-stu-id="af90c-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="af90c-208">1.8</span><span class="sxs-lookup"><span data-stu-id="af90c-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="af90c-209">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-209">String</span></span> | <span data-ttu-id="af90c-210">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="af90c-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="af90c-211">1.5</span><span class="sxs-lookup"><span data-stu-id="af90c-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="af90c-212">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-212">String</span></span> | <span data-ttu-id="af90c-213">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="af90c-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="af90c-214">1.10</span><span class="sxs-lookup"><span data-stu-id="af90c-214">1.10</span></span> |
|`RecipientsChanged`| <span data-ttu-id="af90c-215">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-215">String</span></span> | <span data-ttu-id="af90c-216">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="af90c-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="af90c-217">1.7</span><span class="sxs-lookup"><span data-stu-id="af90c-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="af90c-218">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-218">String</span></span> | <span data-ttu-id="af90c-219">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="af90c-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="af90c-220">1.7</span><span class="sxs-lookup"><span data-stu-id="af90c-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="af90c-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="af90c-221">Requirements</span></span>

|<span data-ttu-id="af90c-222">要求</span><span class="sxs-lookup"><span data-stu-id="af90c-222">Requirement</span></span>| <span data-ttu-id="af90c-223">值</span><span class="sxs-lookup"><span data-stu-id="af90c-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="af90c-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="af90c-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af90c-225">1.5</span><span class="sxs-lookup"><span data-stu-id="af90c-225">1.5</span></span> |
|[<span data-ttu-id="af90c-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="af90c-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af90c-227">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="af90c-228">SourceProperty：String</span><span class="sxs-lookup"><span data-stu-id="af90c-228">SourceProperty: String</span></span>

<span data-ttu-id="af90c-229">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="af90c-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="af90c-230">类型</span><span class="sxs-lookup"><span data-stu-id="af90c-230">Type</span></span>

*   <span data-ttu-id="af90c-231">String</span><span class="sxs-lookup"><span data-stu-id="af90c-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="af90c-232">属性</span><span class="sxs-lookup"><span data-stu-id="af90c-232">Properties</span></span>

|<span data-ttu-id="af90c-233">名称</span><span class="sxs-lookup"><span data-stu-id="af90c-233">Name</span></span>| <span data-ttu-id="af90c-234">类型</span><span class="sxs-lookup"><span data-stu-id="af90c-234">Type</span></span>| <span data-ttu-id="af90c-235">描述</span><span class="sxs-lookup"><span data-stu-id="af90c-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="af90c-236">字符串</span><span class="sxs-lookup"><span data-stu-id="af90c-236">String</span></span>|<span data-ttu-id="af90c-237">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="af90c-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="af90c-238">String</span><span class="sxs-lookup"><span data-stu-id="af90c-238">String</span></span>|<span data-ttu-id="af90c-239">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="af90c-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af90c-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="af90c-240">Requirements</span></span>

|<span data-ttu-id="af90c-241">要求</span><span class="sxs-lookup"><span data-stu-id="af90c-241">Requirement</span></span>| <span data-ttu-id="af90c-242">值</span><span class="sxs-lookup"><span data-stu-id="af90c-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="af90c-243">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="af90c-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af90c-244">1.1</span><span class="sxs-lookup"><span data-stu-id="af90c-244">1.1</span></span>|
|[<span data-ttu-id="af90c-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="af90c-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af90c-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="af90c-246">Compose or Read</span></span>|
