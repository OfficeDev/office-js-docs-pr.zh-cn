---
title: "\"Context\"-\"邮箱-预览要求集\""
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 8c67f7cf9231dd1c0db0d9a8d4ae9fb48e458435
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629193"
---
# <a name="mailbox"></a><span data-ttu-id="2d6f8-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="2d6f8-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="2d6f8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="2d6f8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="2d6f8-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d6f8-105">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-105">Requirements</span></span>

|<span data-ttu-id="2d6f8-106">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-106">Requirement</span></span>| <span data-ttu-id="2d6f8-107">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-109">1.0</span></span>|
|[<span data-ttu-id="2d6f8-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-111">受限</span><span class="sxs-lookup"><span data-stu-id="2d6f8-111">Restricted</span></span>|
|[<span data-ttu-id="2d6f8-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2d6f8-114">属性</span><span class="sxs-lookup"><span data-stu-id="2d6f8-114">Properties</span></span>

| <span data-ttu-id="2d6f8-115">属性</span><span class="sxs-lookup"><span data-stu-id="2d6f8-115">Property</span></span> | <span data-ttu-id="2d6f8-116">最低</span><span class="sxs-lookup"><span data-stu-id="2d6f8-116">Minimum</span></span><br><span data-ttu-id="2d6f8-117">权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-117">permission level</span></span> | <span data-ttu-id="2d6f8-118">型号</span><span class="sxs-lookup"><span data-stu-id="2d6f8-118">Modes</span></span> | <span data-ttu-id="2d6f8-119">返回类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-119">Return type</span></span> | <span data-ttu-id="2d6f8-120">最低</span><span class="sxs-lookup"><span data-stu-id="2d6f8-120">Minimum</span></span><br><span data-ttu-id="2d6f8-121">要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="2d6f8-122">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="2d6f8-122">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="2d6f8-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-123">ReadItem</span></span> | <span data-ttu-id="2d6f8-124">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-124">Compose</span></span><br><span data-ttu-id="2d6f8-125">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-125">Read</span></span> | <span data-ttu-id="2d6f8-126">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-126">String</span></span> | <span data-ttu-id="2d6f8-127">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-127">1.0</span></span> |
| [<span data-ttu-id="2d6f8-128">masterCategories</span><span class="sxs-lookup"><span data-stu-id="2d6f8-128">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="2d6f8-129">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2d6f8-129">ReadWriteMailbox</span></span> | <span data-ttu-id="2d6f8-130">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-130">Compose</span></span><br><span data-ttu-id="2d6f8-131">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-131">Read</span></span> | [<span data-ttu-id="2d6f8-132">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="2d6f8-132">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories) | <span data-ttu-id="2d6f8-133">预览</span><span class="sxs-lookup"><span data-stu-id="2d6f8-133">Preview</span></span> |
| [<span data-ttu-id="2d6f8-134">restUrl</span><span class="sxs-lookup"><span data-stu-id="2d6f8-134">restUrl</span></span>](#resturl-string) | <span data-ttu-id="2d6f8-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-135">ReadItem</span></span> | <span data-ttu-id="2d6f8-136">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-136">Compose</span></span><br><span data-ttu-id="2d6f8-137">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-137">Read</span></span> | <span data-ttu-id="2d6f8-138">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-138">String</span></span> | <span data-ttu-id="2d6f8-139">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-139">1.5</span></span> |

##### <a name="methods"></a><span data-ttu-id="2d6f8-140">方法</span><span class="sxs-lookup"><span data-stu-id="2d6f8-140">Methods</span></span>

| <span data-ttu-id="2d6f8-141">方法</span><span class="sxs-lookup"><span data-stu-id="2d6f8-141">Method</span></span> | <span data-ttu-id="2d6f8-142">最低</span><span class="sxs-lookup"><span data-stu-id="2d6f8-142">Minimum</span></span><br><span data-ttu-id="2d6f8-143">权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-143">permission level</span></span> | <span data-ttu-id="2d6f8-144">型号</span><span class="sxs-lookup"><span data-stu-id="2d6f8-144">Modes</span></span> | <span data-ttu-id="2d6f8-145">最低</span><span class="sxs-lookup"><span data-stu-id="2d6f8-145">Minimum</span></span><br><span data-ttu-id="2d6f8-146">要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-146">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="2d6f8-147">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2d6f8-147">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="2d6f8-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-148">ReadItem</span></span> | <span data-ttu-id="2d6f8-149">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-149">Compose</span></span><br><span data-ttu-id="2d6f8-150">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-150">Read</span></span> | <span data-ttu-id="2d6f8-151">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-151">1.5</span></span> |
| [<span data-ttu-id="2d6f8-152">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="2d6f8-152">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="2d6f8-153">受限</span><span class="sxs-lookup"><span data-stu-id="2d6f8-153">Restricted</span></span> | <span data-ttu-id="2d6f8-154">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-154">Compose</span></span><br><span data-ttu-id="2d6f8-155">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-155">Read</span></span> | <span data-ttu-id="2d6f8-156">1.3</span><span class="sxs-lookup"><span data-stu-id="2d6f8-156">1.3</span></span> |
| [<span data-ttu-id="2d6f8-157">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="2d6f8-157">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="2d6f8-158">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-158">ReadItem</span></span> | <span data-ttu-id="2d6f8-159">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-159">Compose</span></span><br><span data-ttu-id="2d6f8-160">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-160">Read</span></span> | <span data-ttu-id="2d6f8-161">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-161">1.0</span></span> |
| [<span data-ttu-id="2d6f8-162">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="2d6f8-162">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="2d6f8-163">受限</span><span class="sxs-lookup"><span data-stu-id="2d6f8-163">Restricted</span></span> | <span data-ttu-id="2d6f8-164">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-164">Compose</span></span><br><span data-ttu-id="2d6f8-165">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-165">Read</span></span> | <span data-ttu-id="2d6f8-166">1.3</span><span class="sxs-lookup"><span data-stu-id="2d6f8-166">1.3</span></span> |
| [<span data-ttu-id="2d6f8-167">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="2d6f8-167">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="2d6f8-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-168">ReadItem</span></span> | <span data-ttu-id="2d6f8-169">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-169">Compose</span></span><br><span data-ttu-id="2d6f8-170">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-170">Read</span></span> | <span data-ttu-id="2d6f8-171">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-171">1.0</span></span> |
| [<span data-ttu-id="2d6f8-172">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="2d6f8-172">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="2d6f8-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-173">ReadItem</span></span> | <span data-ttu-id="2d6f8-174">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-174">Compose</span></span><br><span data-ttu-id="2d6f8-175">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-175">Read</span></span> | <span data-ttu-id="2d6f8-176">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-176">1.0</span></span> |
| [<span data-ttu-id="2d6f8-177">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="2d6f8-177">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="2d6f8-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-178">ReadItem</span></span> | <span data-ttu-id="2d6f8-179">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-179">Compose</span></span><br><span data-ttu-id="2d6f8-180">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-180">Read</span></span> | <span data-ttu-id="2d6f8-181">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-181">1.0</span></span> |
| [<span data-ttu-id="2d6f8-182">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="2d6f8-182">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="2d6f8-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-183">ReadItem</span></span> | <span data-ttu-id="2d6f8-184">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-184">Read</span></span> | <span data-ttu-id="2d6f8-185">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-185">1.0</span></span> |
| [<span data-ttu-id="2d6f8-186">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="2d6f8-186">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="2d6f8-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-187">ReadItem</span></span> | <span data-ttu-id="2d6f8-188">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-188">Compose</span></span><br><span data-ttu-id="2d6f8-189">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-189">Read</span></span> | <span data-ttu-id="2d6f8-190">1.6</span><span class="sxs-lookup"><span data-stu-id="2d6f8-190">1.6</span></span> |
| [<span data-ttu-id="2d6f8-191">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2d6f8-191">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="2d6f8-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-192">ReadItem</span></span> | <span data-ttu-id="2d6f8-193">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-193">Compose</span></span><br><span data-ttu-id="2d6f8-194">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-194">Read</span></span> | <span data-ttu-id="2d6f8-195">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-195">1.5</span></span> |
| [<span data-ttu-id="2d6f8-196">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2d6f8-196">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="2d6f8-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-197">ReadItem</span></span> | <span data-ttu-id="2d6f8-198">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-198">Compose</span></span><br><span data-ttu-id="2d6f8-199">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-199">Read</span></span> | <span data-ttu-id="2d6f8-200">1.3</span><span class="sxs-lookup"><span data-stu-id="2d6f8-200">1.3</span></span><br><span data-ttu-id="2d6f8-201">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-201">1.0</span></span> |
| [<span data-ttu-id="2d6f8-202">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2d6f8-202">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="2d6f8-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-203">ReadItem</span></span> | <span data-ttu-id="2d6f8-204">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-204">Compose</span></span><br><span data-ttu-id="2d6f8-205">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-205">Read</span></span> | <span data-ttu-id="2d6f8-206">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-206">1.0</span></span> |
| [<span data-ttu-id="2d6f8-207">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="2d6f8-207">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="2d6f8-208">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2d6f8-208">ReadWriteMailbox</span></span> | <span data-ttu-id="2d6f8-209">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-209">Compose</span></span><br><span data-ttu-id="2d6f8-210">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-210">Read</span></span> | <span data-ttu-id="2d6f8-211">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-211">1.0</span></span> |
| [<span data-ttu-id="2d6f8-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2d6f8-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="2d6f8-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-213">ReadItem</span></span> | <span data-ttu-id="2d6f8-214">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-214">Compose</span></span><br><span data-ttu-id="2d6f8-215">读取</span><span class="sxs-lookup"><span data-stu-id="2d6f8-215">Read</span></span> | <span data-ttu-id="2d6f8-216">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-216">1.5</span></span> |

##### <a name="events"></a><span data-ttu-id="2d6f8-217">活动</span><span class="sxs-lookup"><span data-stu-id="2d6f8-217">Events</span></span>

<span data-ttu-id="2d6f8-218">您可以分别使用[addHandlerAsync](#addhandlerasynceventtype-handler-options-callback)和[removeHandlerAsync](#removehandlerasynceventtype-options-callback)订阅和取消订阅以下事件。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-218">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="2d6f8-219">事件</span><span class="sxs-lookup"><span data-stu-id="2d6f8-219">Event</span></span> | <span data-ttu-id="2d6f8-220">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-220">Description</span></span> | <span data-ttu-id="2d6f8-221">最低</span><span class="sxs-lookup"><span data-stu-id="2d6f8-221">Minimum</span></span><br><span data-ttu-id="2d6f8-222">要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-222">requirement set</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="2d6f8-223">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-223">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2d6f8-224">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-224">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="2d6f8-225">邮箱上的 Office 主题已更改。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-225">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="2d6f8-226">预览</span><span class="sxs-lookup"><span data-stu-id="2d6f8-226">Preview</span></span> |

### <a name="namespaces"></a><span data-ttu-id="2d6f8-227">命名空间</span><span class="sxs-lookup"><span data-stu-id="2d6f8-227">Namespaces</span></span>

<span data-ttu-id="2d6f8-228">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-228">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="2d6f8-229">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-229">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="2d6f8-230">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-230">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

## <a name="property-details"></a><span data-ttu-id="2d6f8-231">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="2d6f8-231">Property details</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="2d6f8-232">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-232">ewsUrl: String</span></span>

<span data-ttu-id="2d6f8-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-235">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-235">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2d6f8-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="2d6f8-238">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-238">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="2d6f8-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="2d6f8-241">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-241">Type</span></span>

*   <span data-ttu-id="2d6f8-242">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-242">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d6f8-243">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-243">Requirements</span></span>

|<span data-ttu-id="2d6f8-244">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-244">Requirement</span></span>| <span data-ttu-id="2d6f8-245">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-246">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-247">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-247">1.0</span></span>|
|[<span data-ttu-id="2d6f8-248">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-248">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-249">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-250">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-250">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-251">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-251">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="2d6f8-252">masterCategories： [masterCategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="2d6f8-252">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="2d6f8-253">获取一个对象，该对象提供用于管理此邮箱上的类别主列表的方法。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-253">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-254">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-254">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2d6f8-255">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-255">Type</span></span>

*   [<span data-ttu-id="2d6f8-256">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="2d6f8-256">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="2d6f8-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-257">Requirements</span></span>

|<span data-ttu-id="2d6f8-258">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-258">Requirement</span></span>| <span data-ttu-id="2d6f8-259">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-260">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-261">1.8</span><span class="sxs-lookup"><span data-stu-id="2d6f8-261">1.8</span></span> |
|[<span data-ttu-id="2d6f8-262">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-262">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-263">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2d6f8-263">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="2d6f8-264">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-264">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-265">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-265">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="2d6f8-266">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-266">Example</span></span>

<span data-ttu-id="2d6f8-267">本示例获取此邮箱的类别主列表。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-267">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="2d6f8-268">restUrl：String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-268">restUrl: String</span></span>

<span data-ttu-id="2d6f8-269">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-269">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="2d6f8-270">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-270">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="2d6f8-271">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-271">Type</span></span>

*   <span data-ttu-id="2d6f8-272">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d6f8-273">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-273">Requirements</span></span>

|<span data-ttu-id="2d6f8-274">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-274">Requirement</span></span>| <span data-ttu-id="2d6f8-275">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-276">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-277">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-277">1.5</span></span> |
|[<span data-ttu-id="2d6f8-278">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-279">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-280">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-281">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-281">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="2d6f8-282">方法详细信息</span><span class="sxs-lookup"><span data-stu-id="2d6f8-282">Method details</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="2d6f8-283">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2d6f8-283">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="2d6f8-284">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-284">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="2d6f8-285">目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-285">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-286">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-286">Parameters</span></span>

| <span data-ttu-id="2d6f8-287">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-287">Name</span></span> | <span data-ttu-id="2d6f8-288">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-288">Type</span></span> | <span data-ttu-id="2d6f8-289">属性</span><span class="sxs-lookup"><span data-stu-id="2d6f8-289">Attributes</span></span> | <span data-ttu-id="2d6f8-290">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-290">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2d6f8-291">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2d6f8-291">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2d6f8-292">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-292">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="2d6f8-293">函数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-293">Function</span></span> || <span data-ttu-id="2d6f8-p104">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="2d6f8-297">Object</span><span class="sxs-lookup"><span data-stu-id="2d6f8-297">Object</span></span> | <span data-ttu-id="2d6f8-298">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-298">&lt;optional&gt;</span></span> | <span data-ttu-id="2d6f8-299">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-299">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2d6f8-300">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-300">Object</span></span> | <span data-ttu-id="2d6f8-301">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-301">&lt;optional&gt;</span></span> | <span data-ttu-id="2d6f8-302">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-302">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2d6f8-303">函数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-303">function</span></span>| <span data-ttu-id="2d6f8-304">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-304">&lt;optional&gt;</span></span>|<span data-ttu-id="2d6f8-305">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-305">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-306">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-306">Requirements</span></span>

|<span data-ttu-id="2d6f8-307">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-307">Requirement</span></span>| <span data-ttu-id="2d6f8-308">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-309">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-310">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-310">1.5</span></span> |
|[<span data-ttu-id="2d6f8-311">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-312">ReadItem</span></span> |
|[<span data-ttu-id="2d6f8-313">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-314">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-314">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d6f8-315">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-315">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="2d6f8-316">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="2d6f8-316">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="2d6f8-317">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-317">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-318">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-318">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2d6f8-p105">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-321">参数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-321">Parameters</span></span>

|<span data-ttu-id="2d6f8-322">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-322">Name</span></span>| <span data-ttu-id="2d6f8-323">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-323">Type</span></span>| <span data-ttu-id="2d6f8-324">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-324">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2d6f8-325">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-325">String</span></span>|<span data-ttu-id="2d6f8-326">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-326">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="2d6f8-327">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="2d6f8-327">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="2d6f8-328">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-328">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-329">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-329">Requirements</span></span>

|<span data-ttu-id="2d6f8-330">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-330">Requirement</span></span>| <span data-ttu-id="2d6f8-331">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-331">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-332">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-332">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-333">1.3</span><span class="sxs-lookup"><span data-stu-id="2d6f8-333">1.3</span></span>|
|[<span data-ttu-id="2d6f8-334">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-334">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-335">受限</span><span class="sxs-lookup"><span data-stu-id="2d6f8-335">Restricted</span></span>|
|[<span data-ttu-id="2d6f8-336">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-336">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-337">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-337">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2d6f8-338">返回：</span><span class="sxs-lookup"><span data-stu-id="2d6f8-338">Returns:</span></span>

<span data-ttu-id="2d6f8-339">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="2d6f8-339">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="2d6f8-340">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-340">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="2d6f8-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="2d6f8-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="2d6f8-342">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-342">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="2d6f8-p106">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="2d6f8-p107">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-348">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-348">Parameters</span></span>

|<span data-ttu-id="2d6f8-349">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-349">Name</span></span>| <span data-ttu-id="2d6f8-350">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-350">Type</span></span>| <span data-ttu-id="2d6f8-351">描述</span><span class="sxs-lookup"><span data-stu-id="2d6f8-351">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="2d6f8-352">日期</span><span class="sxs-lookup"><span data-stu-id="2d6f8-352">Date</span></span>|<span data-ttu-id="2d6f8-353">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-353">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-354">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-354">Requirements</span></span>

|<span data-ttu-id="2d6f8-355">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-355">Requirement</span></span>| <span data-ttu-id="2d6f8-356">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-357">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-358">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-358">1.0</span></span>|
|[<span data-ttu-id="2d6f8-359">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-360">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-361">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-362">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-362">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2d6f8-363">返回：</span><span class="sxs-lookup"><span data-stu-id="2d6f8-363">Returns:</span></span>

<span data-ttu-id="2d6f8-364">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="2d6f8-364">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="2d6f8-365">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="2d6f8-365">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="2d6f8-366">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-366">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-367">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-367">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2d6f8-p108">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-370">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-370">Parameters</span></span>

|<span data-ttu-id="2d6f8-371">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-371">Name</span></span>| <span data-ttu-id="2d6f8-372">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-372">Type</span></span>| <span data-ttu-id="2d6f8-373">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-373">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2d6f8-374">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-374">String</span></span>|<span data-ttu-id="2d6f8-375">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-375">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="2d6f8-376">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="2d6f8-376">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="2d6f8-377">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-377">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-378">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-378">Requirements</span></span>

|<span data-ttu-id="2d6f8-379">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-379">Requirement</span></span>| <span data-ttu-id="2d6f8-380">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-381">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-382">1.3</span><span class="sxs-lookup"><span data-stu-id="2d6f8-382">1.3</span></span>|
|[<span data-ttu-id="2d6f8-383">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-384">受限</span><span class="sxs-lookup"><span data-stu-id="2d6f8-384">Restricted</span></span>|
|[<span data-ttu-id="2d6f8-385">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-386">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-386">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2d6f8-387">返回：</span><span class="sxs-lookup"><span data-stu-id="2d6f8-387">Returns:</span></span>

<span data-ttu-id="2d6f8-388">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="2d6f8-388">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="2d6f8-389">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-389">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="2d6f8-390">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="2d6f8-390">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="2d6f8-391">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-391">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="2d6f8-392">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-392">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-393">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-393">Parameters</span></span>

|<span data-ttu-id="2d6f8-394">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-394">Name</span></span>| <span data-ttu-id="2d6f8-395">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-395">Type</span></span>| <span data-ttu-id="2d6f8-396">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-396">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="2d6f8-397">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="2d6f8-397">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="2d6f8-398">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-398">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-399">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-399">Requirements</span></span>

|<span data-ttu-id="2d6f8-400">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-400">Requirement</span></span>| <span data-ttu-id="2d6f8-401">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-402">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-403">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-403">1.0</span></span>|
|[<span data-ttu-id="2d6f8-404">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-405">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-406">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-407">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-407">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2d6f8-408">返回：</span><span class="sxs-lookup"><span data-stu-id="2d6f8-408">Returns:</span></span>

<span data-ttu-id="2d6f8-409">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-409">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="2d6f8-410">键入：日期</span><span class="sxs-lookup"><span data-stu-id="2d6f8-410">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="2d6f8-411">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-411">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="2d6f8-412">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2d6f8-412">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="2d6f8-413">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-413">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-414">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-414">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2d6f8-415">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-415">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2d6f8-p109">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="2d6f8-418">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-418">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="2d6f8-419">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-419">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-420">参数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-420">Parameters</span></span>

|<span data-ttu-id="2d6f8-421">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-421">Name</span></span>| <span data-ttu-id="2d6f8-422">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-422">Type</span></span>| <span data-ttu-id="2d6f8-423">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-423">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2d6f8-424">字符串</span><span class="sxs-lookup"><span data-stu-id="2d6f8-424">String</span></span>|<span data-ttu-id="2d6f8-425">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-425">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-426">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-426">Requirements</span></span>

|<span data-ttu-id="2d6f8-427">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-427">Requirement</span></span>| <span data-ttu-id="2d6f8-428">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-429">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-430">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-430">1.0</span></span>|
|[<span data-ttu-id="2d6f8-431">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-432">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-433">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-434">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d6f8-435">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-435">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="2d6f8-436">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2d6f8-436">displayMessageForm(itemId)</span></span>

<span data-ttu-id="2d6f8-437">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-437">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-438">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-438">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2d6f8-439">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-439">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2d6f8-440">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-440">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="2d6f8-441">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-441">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="2d6f8-p110">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-444">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-444">Parameters</span></span>

|<span data-ttu-id="2d6f8-445">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-445">Name</span></span>| <span data-ttu-id="2d6f8-446">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-446">Type</span></span>| <span data-ttu-id="2d6f8-447">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-447">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2d6f8-448">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-448">String</span></span>|<span data-ttu-id="2d6f8-449">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-449">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-450">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-450">Requirements</span></span>

|<span data-ttu-id="2d6f8-451">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-451">Requirement</span></span>| <span data-ttu-id="2d6f8-452">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-453">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-454">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-454">1.0</span></span>|
|[<span data-ttu-id="2d6f8-455">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-456">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-457">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-458">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-458">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d6f8-459">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-459">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="2d6f8-460">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="2d6f8-460">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="2d6f8-461">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-461">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-462">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-462">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2d6f8-p111">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="2d6f8-p112">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="2d6f8-p113">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="2d6f8-470">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-470">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-471">参数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-471">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-472">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-472">All parameters are optional.</span></span>

|<span data-ttu-id="2d6f8-473">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-473">Name</span></span>| <span data-ttu-id="2d6f8-474">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-474">Type</span></span>| <span data-ttu-id="2d6f8-475">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-475">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="2d6f8-476">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-476">Object</span></span> | <span data-ttu-id="2d6f8-477">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-477">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="2d6f8-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2d6f8-p114">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="2d6f8-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2d6f8-p115">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="2d6f8-484">日期</span><span class="sxs-lookup"><span data-stu-id="2d6f8-484">Date</span></span> | <span data-ttu-id="2d6f8-485">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-485">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="2d6f8-486">Date</span><span class="sxs-lookup"><span data-stu-id="2d6f8-486">Date</span></span> | <span data-ttu-id="2d6f8-487">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-487">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="2d6f8-488">字符串</span><span class="sxs-lookup"><span data-stu-id="2d6f8-488">String</span></span> | <span data-ttu-id="2d6f8-p116">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="2d6f8-491">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-491">Array.&lt;String&gt;</span></span> | <span data-ttu-id="2d6f8-p117">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="2d6f8-494">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-494">String</span></span> | <span data-ttu-id="2d6f8-p118">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="2d6f8-497">字符串</span><span class="sxs-lookup"><span data-stu-id="2d6f8-497">String</span></span> | <span data-ttu-id="2d6f8-p119">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2d6f8-500">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-500">Requirements</span></span>

|<span data-ttu-id="2d6f8-501">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-501">Requirement</span></span>| <span data-ttu-id="2d6f8-502">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-503">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-504">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-504">1.0</span></span>|
|[<span data-ttu-id="2d6f8-505">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-506">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-507">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-508">阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d6f8-509">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-509">Example</span></span>

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="2d6f8-510">Office.context.mailbox.displaynewmessageform （参数）</span><span class="sxs-lookup"><span data-stu-id="2d6f8-510">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="2d6f8-511">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-511">Displays a form for creating a new message.</span></span>

<span data-ttu-id="2d6f8-512">`displayNewMessageForm`方法将打开一个窗体，使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-512">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="2d6f8-513">如果指定了参数，则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-513">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="2d6f8-514">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-514">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-515">参数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-515">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-516">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-516">All parameters are optional.</span></span>

|<span data-ttu-id="2d6f8-517">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-517">Name</span></span>| <span data-ttu-id="2d6f8-518">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-518">Type</span></span>| <span data-ttu-id="2d6f8-519">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-519">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="2d6f8-520">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-520">Object</span></span> | <span data-ttu-id="2d6f8-521">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-521">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="2d6f8-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2d6f8-523">包含电子邮件地址的字符串数组，或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-523">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="2d6f8-524">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-524">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="2d6f8-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2d6f8-526">包含电子邮件地址的字符串数组，或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-526">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="2d6f8-527">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-527">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="2d6f8-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2d6f8-529">包含电子邮件地址的字符串数组，或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-529">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="2d6f8-530">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-530">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="2d6f8-531">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-531">String</span></span> | <span data-ttu-id="2d6f8-532">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-532">A string containing the subject of the message.</span></span> <span data-ttu-id="2d6f8-533">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-533">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="2d6f8-534">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-534">String</span></span> | <span data-ttu-id="2d6f8-535">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-535">The HTML body of the message.</span></span> <span data-ttu-id="2d6f8-536">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-536">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="2d6f8-537">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-537">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2d6f8-538">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-538">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="2d6f8-539">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-539">String</span></span> | <span data-ttu-id="2d6f8-p126">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="2d6f8-542">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-542">String</span></span> | <span data-ttu-id="2d6f8-543">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-543">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="2d6f8-544">String</span><span class="sxs-lookup"><span data-stu-id="2d6f8-544">String</span></span> | <span data-ttu-id="2d6f8-p127">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="2d6f8-547">布尔</span><span class="sxs-lookup"><span data-stu-id="2d6f8-547">Boolean</span></span> | <span data-ttu-id="2d6f8-p128">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="2d6f8-550">字符串</span><span class="sxs-lookup"><span data-stu-id="2d6f8-550">String</span></span> | <span data-ttu-id="2d6f8-551">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-551">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="2d6f8-552">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-552">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="2d6f8-553">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-553">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="2d6f8-554">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-554">Requirements</span></span>

|<span data-ttu-id="2d6f8-555">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-555">Requirement</span></span>| <span data-ttu-id="2d6f8-556">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-557">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-558">1.6</span><span class="sxs-lookup"><span data-stu-id="2d6f8-558">1.6</span></span> |
|[<span data-ttu-id="2d6f8-559">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-560">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-561">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-562">阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-562">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d6f8-563">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-563">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="2d6f8-564">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2d6f8-564">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="2d6f8-565">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-565">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="2d6f8-p130">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-568">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-568">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="2d6f8-569">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-569">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="2d6f8-570">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-570">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="2d6f8-571">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-571">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="2d6f8-572">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="2d6f8-572">**REST Tokens**</span></span>

<span data-ttu-id="2d6f8-p132">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="2d6f8-576">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-576">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="2d6f8-577">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="2d6f8-577">**EWS Tokens**</span></span>

<span data-ttu-id="2d6f8-p133">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="2d6f8-580">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-580">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="2d6f8-581">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-581">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="2d6f8-582">第三方系统使用令牌作为持有者授权令牌，以调用 Exchange Web 服务（EWS） [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作或[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作以检索附件或项目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-582">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="2d6f8-583">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-583">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-584">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-584">Parameters</span></span>

|<span data-ttu-id="2d6f8-585">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-585">Name</span></span>| <span data-ttu-id="2d6f8-586">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-586">Type</span></span>| <span data-ttu-id="2d6f8-587">属性</span><span class="sxs-lookup"><span data-stu-id="2d6f8-587">Attributes</span></span>| <span data-ttu-id="2d6f8-588">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-588">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="2d6f8-589">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-589">Object</span></span> | <span data-ttu-id="2d6f8-590">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-590">&lt;optional&gt;</span></span> | <span data-ttu-id="2d6f8-591">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-591">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="2d6f8-592">布尔值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-592">Boolean</span></span> |  <span data-ttu-id="2d6f8-593">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-593">&lt;optional&gt;</span></span> | <span data-ttu-id="2d6f8-p135">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2d6f8-596">Object</span><span class="sxs-lookup"><span data-stu-id="2d6f8-596">Object</span></span> |  <span data-ttu-id="2d6f8-597">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-597">&lt;optional&gt;</span></span> | <span data-ttu-id="2d6f8-598">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-598">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="2d6f8-599">函数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-599">function</span></span>||<span data-ttu-id="2d6f8-600">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-600">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2d6f8-601">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-601">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="2d6f8-602">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-602">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2d6f8-603">错误</span><span class="sxs-lookup"><span data-stu-id="2d6f8-603">Errors</span></span>

|<span data-ttu-id="2d6f8-604">错误代码</span><span class="sxs-lookup"><span data-stu-id="2d6f8-604">Error code</span></span>|<span data-ttu-id="2d6f8-605">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-605">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="2d6f8-606">请求失败。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-606">The request has failed.</span></span> <span data-ttu-id="2d6f8-607">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-607">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="2d6f8-608">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-608">The Exchange server returned an error.</span></span> <span data-ttu-id="2d6f8-609">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-609">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="2d6f8-610">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-610">The user is no longer connected to the network.</span></span> <span data-ttu-id="2d6f8-611">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-611">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-612">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-612">Requirements</span></span>

|<span data-ttu-id="2d6f8-613">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-613">Requirement</span></span>| <span data-ttu-id="2d6f8-614">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-615">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-616">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-616">1.5</span></span> |
|[<span data-ttu-id="2d6f8-617">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-618">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-619">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-620">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-620">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d6f8-621">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-621">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="2d6f8-622">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2d6f8-622">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2d6f8-623">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-623">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="2d6f8-p139">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="2d6f8-626">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-626">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="2d6f8-627">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-627">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="2d6f8-628">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-628">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="2d6f8-629">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-629">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="2d6f8-630">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-630">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="2d6f8-631">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-631">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-632">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-632">Parameters</span></span>

|<span data-ttu-id="2d6f8-633">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-633">Name</span></span>| <span data-ttu-id="2d6f8-634">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-634">Type</span></span>| <span data-ttu-id="2d6f8-635">属性</span><span class="sxs-lookup"><span data-stu-id="2d6f8-635">Attributes</span></span>| <span data-ttu-id="2d6f8-636">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-636">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2d6f8-637">function</span><span class="sxs-lookup"><span data-stu-id="2d6f8-637">function</span></span>||<span data-ttu-id="2d6f8-638">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2d6f8-639">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-639">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="2d6f8-640">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-640">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="2d6f8-641">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-641">Object</span></span>| <span data-ttu-id="2d6f8-642">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-642">&lt;optional&gt;</span></span>|<span data-ttu-id="2d6f8-643">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-643">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2d6f8-644">错误</span><span class="sxs-lookup"><span data-stu-id="2d6f8-644">Errors</span></span>

|<span data-ttu-id="2d6f8-645">错误代码</span><span class="sxs-lookup"><span data-stu-id="2d6f8-645">Error code</span></span>|<span data-ttu-id="2d6f8-646">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-646">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="2d6f8-647">请求失败。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-647">The request has failed.</span></span> <span data-ttu-id="2d6f8-648">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-648">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="2d6f8-649">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-649">The Exchange server returned an error.</span></span> <span data-ttu-id="2d6f8-650">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-650">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="2d6f8-651">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-651">The user is no longer connected to the network.</span></span> <span data-ttu-id="2d6f8-652">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-652">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-653">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-653">Requirements</span></span>

|<span data-ttu-id="2d6f8-654">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-654">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2d6f8-655">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-656">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-656">1.0</span></span> | <span data-ttu-id="2d6f8-657">1.3</span><span class="sxs-lookup"><span data-stu-id="2d6f8-657">1.3</span></span> |
|[<span data-ttu-id="2d6f8-658">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-658">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-659">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-659">ReadItem</span></span> | <span data-ttu-id="2d6f8-660">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-660">ReadItem</span></span> |
|[<span data-ttu-id="2d6f8-661">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-661">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-662">阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-662">Read</span></span> | <span data-ttu-id="2d6f8-663">撰写</span><span class="sxs-lookup"><span data-stu-id="2d6f8-663">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="2d6f8-664">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-664">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="2d6f8-665">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2d6f8-665">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2d6f8-666">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-666">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="2d6f8-667">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-667">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-668">参数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-668">Parameters</span></span>

|<span data-ttu-id="2d6f8-669">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-669">Name</span></span>| <span data-ttu-id="2d6f8-670">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-670">Type</span></span>| <span data-ttu-id="2d6f8-671">属性</span><span class="sxs-lookup"><span data-stu-id="2d6f8-671">Attributes</span></span>| <span data-ttu-id="2d6f8-672">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-672">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2d6f8-673">function</span><span class="sxs-lookup"><span data-stu-id="2d6f8-673">function</span></span>||<span data-ttu-id="2d6f8-674">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-674">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2d6f8-675">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-675">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="2d6f8-676">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-676">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="2d6f8-677">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-677">Object</span></span>| <span data-ttu-id="2d6f8-678">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-678">&lt;optional&gt;</span></span>|<span data-ttu-id="2d6f8-679">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-679">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2d6f8-680">错误</span><span class="sxs-lookup"><span data-stu-id="2d6f8-680">Errors</span></span>

|<span data-ttu-id="2d6f8-681">错误代码</span><span class="sxs-lookup"><span data-stu-id="2d6f8-681">Error code</span></span>|<span data-ttu-id="2d6f8-682">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-682">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="2d6f8-683">请求失败。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-683">The request has failed.</span></span> <span data-ttu-id="2d6f8-684">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-684">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="2d6f8-685">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-685">The Exchange server returned an error.</span></span> <span data-ttu-id="2d6f8-686">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-686">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="2d6f8-687">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-687">The user is no longer connected to the network.</span></span> <span data-ttu-id="2d6f8-688">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-688">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-689">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-689">Requirements</span></span>

|<span data-ttu-id="2d6f8-690">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-690">Requirement</span></span>| <span data-ttu-id="2d6f8-691">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-691">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-692">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-692">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-693">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-693">1.0</span></span>|
|[<span data-ttu-id="2d6f8-694">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-694">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-695">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-695">ReadItem</span></span>|
|[<span data-ttu-id="2d6f8-696">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-696">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-697">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-697">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d6f8-698">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-698">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="2d6f8-699">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2d6f8-699">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="2d6f8-700">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-700">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-701">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-701">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="2d6f8-702">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="2d6f8-702">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="2d6f8-703">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="2d6f8-703">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="2d6f8-704">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-704">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="2d6f8-705">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-705">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="2d6f8-706">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-706">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="2d6f8-707">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-707">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="2d6f8-708">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-708">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="2d6f8-p149">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="2d6f8-711">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-711">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="2d6f8-712">版本差异</span><span class="sxs-lookup"><span data-stu-id="2d6f8-712">Version differences</span></span>

<span data-ttu-id="2d6f8-713">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-713">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="2d6f8-714">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-714">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="2d6f8-715">您可以使用邮箱. hostName 属性确定您的邮件应用程序是在 web 上的 Outlook 中运行还是在桌面客户端上运行。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-715">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="2d6f8-716">可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-716">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-717">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-717">Parameters</span></span>

|<span data-ttu-id="2d6f8-718">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-718">Name</span></span>| <span data-ttu-id="2d6f8-719">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-719">Type</span></span>| <span data-ttu-id="2d6f8-720">属性</span><span class="sxs-lookup"><span data-stu-id="2d6f8-720">Attributes</span></span>| <span data-ttu-id="2d6f8-721">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-721">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="2d6f8-722">字符串</span><span class="sxs-lookup"><span data-stu-id="2d6f8-722">String</span></span>||<span data-ttu-id="2d6f8-723">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-723">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="2d6f8-724">函数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-724">function</span></span>||<span data-ttu-id="2d6f8-725">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2d6f8-726">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-726">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="2d6f8-727">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-727">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="2d6f8-728">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-728">Object</span></span>| <span data-ttu-id="2d6f8-729">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-729">&lt;optional&gt;</span></span>|<span data-ttu-id="2d6f8-730">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-730">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-731">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-731">Requirements</span></span>

|<span data-ttu-id="2d6f8-732">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-732">Requirement</span></span>| <span data-ttu-id="2d6f8-733">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-734">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-735">1.0</span><span class="sxs-lookup"><span data-stu-id="2d6f8-735">1.0</span></span>|
|[<span data-ttu-id="2d6f8-736">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-737">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2d6f8-737">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="2d6f8-738">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-739">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-739">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d6f8-740">示例</span><span class="sxs-lookup"><span data-stu-id="2d6f8-740">Example</span></span>

<span data-ttu-id="2d6f8-741">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-741">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="2d6f8-742">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2d6f8-742">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="2d6f8-743">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-743">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="2d6f8-744">目前，支持的事件类型为`Office.EventType.ItemChanged`和`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-744">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2d6f8-745">Parameters</span><span class="sxs-lookup"><span data-stu-id="2d6f8-745">Parameters</span></span>

| <span data-ttu-id="2d6f8-746">名称</span><span class="sxs-lookup"><span data-stu-id="2d6f8-746">Name</span></span> | <span data-ttu-id="2d6f8-747">类型</span><span class="sxs-lookup"><span data-stu-id="2d6f8-747">Type</span></span> | <span data-ttu-id="2d6f8-748">属性</span><span class="sxs-lookup"><span data-stu-id="2d6f8-748">Attributes</span></span> | <span data-ttu-id="2d6f8-749">说明</span><span class="sxs-lookup"><span data-stu-id="2d6f8-749">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2d6f8-750">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2d6f8-750">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2d6f8-751">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-751">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="2d6f8-752">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-752">Object</span></span> | <span data-ttu-id="2d6f8-753">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-753">&lt;optional&gt;</span></span> | <span data-ttu-id="2d6f8-754">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-754">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2d6f8-755">对象</span><span class="sxs-lookup"><span data-stu-id="2d6f8-755">Object</span></span> | <span data-ttu-id="2d6f8-756">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-756">&lt;optional&gt;</span></span> | <span data-ttu-id="2d6f8-757">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-757">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2d6f8-758">函数</span><span class="sxs-lookup"><span data-stu-id="2d6f8-758">function</span></span>| <span data-ttu-id="2d6f8-759">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2d6f8-759">&lt;optional&gt;</span></span>|<span data-ttu-id="2d6f8-760">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2d6f8-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d6f8-761">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d6f8-761">Requirements</span></span>

|<span data-ttu-id="2d6f8-762">要求</span><span class="sxs-lookup"><span data-stu-id="2d6f8-762">Requirement</span></span>| <span data-ttu-id="2d6f8-763">值</span><span class="sxs-lookup"><span data-stu-id="2d6f8-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d6f8-764">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2d6f8-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d6f8-765">1.5</span><span class="sxs-lookup"><span data-stu-id="2d6f8-765">1.5</span></span> |
|[<span data-ttu-id="2d6f8-766">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2d6f8-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d6f8-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d6f8-767">ReadItem</span></span> |
|[<span data-ttu-id="2d6f8-768">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2d6f8-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d6f8-769">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2d6f8-769">Compose or Read</span></span>|
