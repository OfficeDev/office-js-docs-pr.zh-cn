---
title: "\"Context.subname\"-\"邮箱-要求集 1.7\""
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 95fb4ce6bcc3c44c77dc4623a12b140ca979949c
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127147"
---
# <a name="mailbox"></a><span data-ttu-id="e27a7-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="e27a7-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="e27a7-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="e27a7-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="e27a7-104">提供对 Microsoft Outlook 的 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="e27a7-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e27a7-105">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-105">Requirements</span></span>

|<span data-ttu-id="e27a7-106">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-106">Requirement</span></span>| <span data-ttu-id="e27a7-107">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-109">1.0</span></span>|
|[<span data-ttu-id="e27a7-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-111">受限</span><span class="sxs-lookup"><span data-stu-id="e27a7-111">Restricted</span></span>|
|[<span data-ttu-id="e27a7-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e27a7-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-114">Members and methods</span></span>

| <span data-ttu-id="e27a7-115">成员</span><span class="sxs-lookup"><span data-stu-id="e27a7-115">Member</span></span> | <span data-ttu-id="e27a7-116">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e27a7-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="e27a7-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="e27a7-118">成员</span><span class="sxs-lookup"><span data-stu-id="e27a7-118">Member</span></span> |
| [<span data-ttu-id="e27a7-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="e27a7-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="e27a7-120">成员</span><span class="sxs-lookup"><span data-stu-id="e27a7-120">Member</span></span> |
| [<span data-ttu-id="e27a7-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e27a7-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e27a7-122">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-122">Method</span></span> |
| [<span data-ttu-id="e27a7-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="e27a7-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="e27a7-124">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-124">Method</span></span> |
| [<span data-ttu-id="e27a7-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e27a7-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="e27a7-126">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-126">Method</span></span> |
| [<span data-ttu-id="e27a7-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="e27a7-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="e27a7-128">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-128">Method</span></span> |
| [<span data-ttu-id="e27a7-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="e27a7-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="e27a7-130">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-130">Method</span></span> |
| [<span data-ttu-id="e27a7-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e27a7-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="e27a7-132">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-132">Method</span></span> |
| [<span data-ttu-id="e27a7-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="e27a7-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="e27a7-134">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-134">Method</span></span> |
| [<span data-ttu-id="e27a7-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e27a7-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="e27a7-136">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-136">Method</span></span> |
| [<span data-ttu-id="e27a7-137">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="e27a7-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="e27a7-138">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-138">Method</span></span> |
| [<span data-ttu-id="e27a7-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e27a7-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="e27a7-140">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-140">Method</span></span> |
| [<span data-ttu-id="e27a7-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e27a7-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="e27a7-142">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-142">Method</span></span> |
| [<span data-ttu-id="e27a7-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e27a7-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="e27a7-144">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-144">Method</span></span> |
| [<span data-ttu-id="e27a7-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e27a7-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="e27a7-146">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-146">Method</span></span> |
| [<span data-ttu-id="e27a7-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e27a7-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="e27a7-148">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e27a7-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="e27a7-149">Namespaces</span></span>

<span data-ttu-id="e27a7-150">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="e27a7-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="e27a7-151">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="e27a7-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="e27a7-152">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="e27a7-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="e27a7-153">成员</span><span class="sxs-lookup"><span data-stu-id="e27a7-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="e27a7-154">Mailbox.ewsurl: String</span><span class="sxs-lookup"><span data-stu-id="e27a7-154">ewsUrl: String</span></span>

<span data-ttu-id="e27a7-155">获取此电子邮件帐户的 Exchange Web Services (EWS) 终点的 URL。</span><span class="sxs-lookup"><span data-stu-id="e27a7-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="e27a7-156">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e27a7-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-157">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="e27a7-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e27a7-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e27a7-160">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="e27a7-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="e27a7-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e27a7-163">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-163">Type</span></span>

*   <span data-ttu-id="e27a7-164">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e27a7-165">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-165">Requirements</span></span>

|<span data-ttu-id="e27a7-166">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-166">Requirement</span></span>| <span data-ttu-id="e27a7-167">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-169">1.0</span></span>|
|[<span data-ttu-id="e27a7-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-171">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-173">Compose or Read</span></span>|

---
---

#### <a name="resturl-string"></a><span data-ttu-id="e27a7-174">Office.context.mailbox.resturl: String</span><span class="sxs-lookup"><span data-stu-id="e27a7-174">restUrl: String</span></span>

<span data-ttu-id="e27a7-175">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="e27a7-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="e27a7-176">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="e27a7-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="e27a7-177">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="e27a7-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="e27a7-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e27a7-180">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-180">Type</span></span>

*   <span data-ttu-id="e27a7-181">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e27a7-182">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-182">Requirements</span></span>

|<span data-ttu-id="e27a7-183">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-183">Requirement</span></span>| <span data-ttu-id="e27a7-184">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-185">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-186">1.5</span><span class="sxs-lookup"><span data-stu-id="e27a7-186">1.5</span></span> |
|[<span data-ttu-id="e27a7-187">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-188">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e27a7-191">方法</span><span class="sxs-lookup"><span data-stu-id="e27a7-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e27a7-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e27a7-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e27a7-193">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="e27a7-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e27a7-194">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="e27a7-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-195">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-195">Parameters</span></span>

| <span data-ttu-id="e27a7-196">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-196">Name</span></span> | <span data-ttu-id="e27a7-197">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-197">Type</span></span> | <span data-ttu-id="e27a7-198">属性</span><span class="sxs-lookup"><span data-stu-id="e27a7-198">Attributes</span></span> | <span data-ttu-id="e27a7-199">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e27a7-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e27a7-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e27a7-201">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="e27a7-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e27a7-202">函数</span><span class="sxs-lookup"><span data-stu-id="e27a7-202">Function</span></span> || <span data-ttu-id="e27a7-p105">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e27a7-206">Object</span><span class="sxs-lookup"><span data-stu-id="e27a7-206">Object</span></span> | <span data-ttu-id="e27a7-207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-207">&lt;optional&gt;</span></span> | <span data-ttu-id="e27a7-208">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e27a7-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e27a7-209">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-209">Object</span></span> | <span data-ttu-id="e27a7-210">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-210">&lt;optional&gt;</span></span> | <span data-ttu-id="e27a7-211">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e27a7-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e27a7-212">函数</span><span class="sxs-lookup"><span data-stu-id="e27a7-212">function</span></span>| <span data-ttu-id="e27a7-213">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-213">&lt;optional&gt;</span></span>|<span data-ttu-id="e27a7-214">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e27a7-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-215">Requirements</span><span class="sxs-lookup"><span data-stu-id="e27a7-215">Requirements</span></span>

|<span data-ttu-id="e27a7-216">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-216">Requirement</span></span>| <span data-ttu-id="e27a7-217">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-219">1.5</span><span class="sxs-lookup"><span data-stu-id="e27a7-219">1.5</span></span> |
|[<span data-ttu-id="e27a7-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-221">ReadItem</span></span> |
|[<span data-ttu-id="e27a7-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-223">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-224">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-224">Example</span></span>

```javascript
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
};
```

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="e27a7-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e27a7-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e27a7-226">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="e27a7-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-227">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e27a7-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e27a7-p106">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-230">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-230">Parameters</span></span>

|<span data-ttu-id="e27a7-231">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-231">Name</span></span>| <span data-ttu-id="e27a7-232">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-232">Type</span></span>| <span data-ttu-id="e27a7-233">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e27a7-234">字符串</span><span class="sxs-lookup"><span data-stu-id="e27a7-234">String</span></span>|<span data-ttu-id="e27a7-235">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="e27a7-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="e27a7-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e27a7-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="e27a7-237">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="e27a7-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-238">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-238">Requirements</span></span>

|<span data-ttu-id="e27a7-239">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-239">Requirement</span></span>| <span data-ttu-id="e27a7-240">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-242">1.3</span><span class="sxs-lookup"><span data-stu-id="e27a7-242">1.3</span></span>|
|[<span data-ttu-id="e27a7-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-244">受限</span><span class="sxs-lookup"><span data-stu-id="e27a7-244">Restricted</span></span>|
|[<span data-ttu-id="e27a7-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e27a7-247">返回：</span><span class="sxs-lookup"><span data-stu-id="e27a7-247">Returns:</span></span>

<span data-ttu-id="e27a7-248">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="e27a7-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e27a7-249">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="e27a7-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="e27a7-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="e27a7-251">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="e27a7-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="e27a7-252">适用于桌面或 web 上的 Outlook 的邮件应用程序可以对日期和时间使用不同的时区。</span><span class="sxs-lookup"><span data-stu-id="e27a7-252">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="e27a7-253">桌面上的 Outlook 使用客户端计算机时区;Web 上的 Outlook 使用 Exchange 管理中心 (EAC) 上设置的时区。</span><span class="sxs-lookup"><span data-stu-id="e27a7-253">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="e27a7-254">应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="e27a7-254">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="e27a7-255">如果邮件应用程序在桌面客户端上的 Outlook 中运行, `convertToLocalClientTime`则该方法将返回一个 dictionary 对象, 并将值设置为客户端计算机时区。</span><span class="sxs-lookup"><span data-stu-id="e27a7-255">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="e27a7-256">如果邮件应用程序在 web 上的 Outlook 中运行, 则`convertToLocalClientTime`该方法将返回一个 dictionary 对象, 其中的值设置为 EAC 中指定的时区。</span><span class="sxs-lookup"><span data-stu-id="e27a7-256">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-257">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-257">Parameters</span></span>

|<span data-ttu-id="e27a7-258">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-258">Name</span></span>| <span data-ttu-id="e27a7-259">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-259">Type</span></span>| <span data-ttu-id="e27a7-260">描述</span><span class="sxs-lookup"><span data-stu-id="e27a7-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="e27a7-261">日期</span><span class="sxs-lookup"><span data-stu-id="e27a7-261">Date</span></span>|<span data-ttu-id="e27a7-262">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-263">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-263">Requirements</span></span>

|<span data-ttu-id="e27a7-264">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-264">Requirement</span></span>| <span data-ttu-id="e27a7-265">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-266">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-267">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-267">1.0</span></span>|
|[<span data-ttu-id="e27a7-268">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-269">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-270">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-271">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e27a7-272">返回：</span><span class="sxs-lookup"><span data-stu-id="e27a7-272">Returns:</span></span>

<span data-ttu-id="e27a7-273">类型：[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="e27a7-273">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="e27a7-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e27a7-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e27a7-275">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="e27a7-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-276">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e27a7-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e27a7-p109">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-279">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-279">Parameters</span></span>

|<span data-ttu-id="e27a7-280">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-280">Name</span></span>| <span data-ttu-id="e27a7-281">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-281">Type</span></span>| <span data-ttu-id="e27a7-282">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e27a7-283">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-283">String</span></span>|<span data-ttu-id="e27a7-284">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="e27a7-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="e27a7-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e27a7-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="e27a7-286">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="e27a7-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-287">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-287">Requirements</span></span>

|<span data-ttu-id="e27a7-288">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-288">Requirement</span></span>| <span data-ttu-id="e27a7-289">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-290">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-291">1.3</span><span class="sxs-lookup"><span data-stu-id="e27a7-291">1.3</span></span>|
|[<span data-ttu-id="e27a7-292">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-293">受限</span><span class="sxs-lookup"><span data-stu-id="e27a7-293">Restricted</span></span>|
|[<span data-ttu-id="e27a7-294">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-295">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e27a7-296">返回：</span><span class="sxs-lookup"><span data-stu-id="e27a7-296">Returns:</span></span>

<span data-ttu-id="e27a7-297">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="e27a7-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e27a7-298">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="e27a7-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="e27a7-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="e27a7-300">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="e27a7-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="e27a7-301">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="e27a7-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-302">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-302">Parameters</span></span>

|<span data-ttu-id="e27a7-303">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-303">Name</span></span>| <span data-ttu-id="e27a7-304">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-304">Type</span></span>| <span data-ttu-id="e27a7-305">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="e27a7-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e27a7-306">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="e27a7-307">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="e27a7-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-308">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-308">Requirements</span></span>

|<span data-ttu-id="e27a7-309">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-309">Requirement</span></span>| <span data-ttu-id="e27a7-310">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-312">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-312">1.0</span></span>|
|[<span data-ttu-id="e27a7-313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-314">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-316">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e27a7-317">返回：</span><span class="sxs-lookup"><span data-stu-id="e27a7-317">Returns:</span></span>

<span data-ttu-id="e27a7-318">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="e27a7-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="e27a7-319">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="e27a7-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e27a7-320">日期</span><span class="sxs-lookup"><span data-stu-id="e27a7-320">Date</span></span></dd>

</dl>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="e27a7-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e27a7-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="e27a7-322">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="e27a7-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-323">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e27a7-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e27a7-324">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="e27a7-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e27a7-325">在 Mac 上的 Outlook 中, 可以使用此方法显示不是定期系列的一部分的单个约会, 也可以是定期系列的主约会, 但不能显示该系列的实例。</span><span class="sxs-lookup"><span data-stu-id="e27a7-325">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="e27a7-326">这是因为在 Mac 上的 Outlook 中, 无法访问定期系列的实例的属性 (包括项目 ID)。</span><span class="sxs-lookup"><span data-stu-id="e27a7-326">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="e27a7-327">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于32KB 个字符时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="e27a7-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="e27a7-328">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="e27a7-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-329">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-329">Parameters</span></span>

|<span data-ttu-id="e27a7-330">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-330">Name</span></span>| <span data-ttu-id="e27a7-331">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-331">Type</span></span>| <span data-ttu-id="e27a7-332">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e27a7-333">字符串</span><span class="sxs-lookup"><span data-stu-id="e27a7-333">String</span></span>|<span data-ttu-id="e27a7-334">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="e27a7-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-335">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-335">Requirements</span></span>

|<span data-ttu-id="e27a7-336">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-336">Requirement</span></span>| <span data-ttu-id="e27a7-337">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-338">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-339">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-339">1.0</span></span>|
|[<span data-ttu-id="e27a7-340">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-341">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-342">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-343">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-344">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="e27a7-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e27a7-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="e27a7-346">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="e27a7-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-347">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e27a7-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e27a7-348">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="e27a7-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e27a7-349">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于 32 KB 的字符数时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="e27a7-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="e27a7-350">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="e27a7-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="e27a7-p111">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-353">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-353">Parameters</span></span>

|<span data-ttu-id="e27a7-354">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-354">Name</span></span>| <span data-ttu-id="e27a7-355">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-355">Type</span></span>| <span data-ttu-id="e27a7-356">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e27a7-357">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-357">String</span></span>|<span data-ttu-id="e27a7-358">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="e27a7-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-359">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-359">Requirements</span></span>

|<span data-ttu-id="e27a7-360">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-360">Requirement</span></span>| <span data-ttu-id="e27a7-361">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-363">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-363">1.0</span></span>|
|[<span data-ttu-id="e27a7-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-365">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-368">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="e27a7-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e27a7-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="e27a7-370">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="e27a7-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-371">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e27a7-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e27a7-p112">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e27a7-374">在 web 和移动设备上的 Outlook 中, 此方法始终显示一个包含 "与会者" 字段的窗体。</span><span class="sxs-lookup"><span data-stu-id="e27a7-374">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="e27a7-375">如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。</span><span class="sxs-lookup"><span data-stu-id="e27a7-375">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="e27a7-376">如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="e27a7-376">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="e27a7-p114">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="e27a7-379">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="e27a7-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-380">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-381">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="e27a7-381">All parameters are optional.</span></span>

|<span data-ttu-id="e27a7-382">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-382">Name</span></span>| <span data-ttu-id="e27a7-383">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-383">Type</span></span>| <span data-ttu-id="e27a7-384">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e27a7-385">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-385">Object</span></span> | <span data-ttu-id="e27a7-386">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="e27a7-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="e27a7-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e27a7-p115">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="e27a7-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e27a7-p116">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="e27a7-393">日期</span><span class="sxs-lookup"><span data-stu-id="e27a7-393">Date</span></span> | <span data-ttu-id="e27a7-394">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="e27a7-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="e27a7-395">Date</span><span class="sxs-lookup"><span data-stu-id="e27a7-395">Date</span></span> | <span data-ttu-id="e27a7-396">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="e27a7-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="e27a7-397">字符串</span><span class="sxs-lookup"><span data-stu-id="e27a7-397">String</span></span> | <span data-ttu-id="e27a7-p117">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="e27a7-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="e27a7-p118">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e27a7-403">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-403">String</span></span> | <span data-ttu-id="e27a7-p119">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="e27a7-406">字符串</span><span class="sxs-lookup"><span data-stu-id="e27a7-406">String</span></span> | <span data-ttu-id="e27a7-p120">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e27a7-409">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-409">Requirements</span></span>

|<span data-ttu-id="e27a7-410">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-410">Requirement</span></span>| <span data-ttu-id="e27a7-411">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-412">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-413">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-413">1.0</span></span>|
|[<span data-ttu-id="e27a7-414">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-415">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-416">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-417">阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-418">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-418">Example</span></span>

```javascript
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

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="e27a7-419">Office.context.mailbox.displaynewmessageform (参数)</span><span class="sxs-lookup"><span data-stu-id="e27a7-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="e27a7-420">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="e27a7-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="e27a7-421">`displayNewMessageForm`方法将打开一个窗体, 使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="e27a7-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="e27a7-422">如果指定了参数, 则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="e27a7-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e27a7-423">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="e27a7-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-424">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-425">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="e27a7-425">All parameters are optional.</span></span>

|<span data-ttu-id="e27a7-426">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-426">Name</span></span>| <span data-ttu-id="e27a7-427">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-427">Type</span></span>| <span data-ttu-id="e27a7-428">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e27a7-429">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-429">Object</span></span> | <span data-ttu-id="e27a7-430">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="e27a7-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="e27a7-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e27a7-432">包含电子邮件地址的字符串数组, 或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="e27a7-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="e27a7-433">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="e27a7-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="e27a7-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e27a7-435">包含电子邮件地址的字符串数组, 或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="e27a7-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="e27a7-436">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="e27a7-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="e27a7-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e27a7-438">包含电子邮件地址的字符串数组, 或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="e27a7-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="e27a7-439">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="e27a7-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e27a7-440">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-440">String</span></span> | <span data-ttu-id="e27a7-441">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="e27a7-441">A string containing the subject of the message.</span></span> <span data-ttu-id="e27a7-442">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e27a7-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="e27a7-443">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-443">String</span></span> | <span data-ttu-id="e27a7-444">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="e27a7-444">The HTML body of the message.</span></span> <span data-ttu-id="e27a7-445">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="e27a7-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="e27a7-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e27a7-447">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="e27a7-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="e27a7-448">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-448">String</span></span> | <span data-ttu-id="e27a7-p127">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="e27a7-451">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-451">String</span></span> | <span data-ttu-id="e27a7-452">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e27a7-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="e27a7-453">String</span><span class="sxs-lookup"><span data-stu-id="e27a7-453">String</span></span> | <span data-ttu-id="e27a7-p128">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="e27a7-456">布尔</span><span class="sxs-lookup"><span data-stu-id="e27a7-456">Boolean</span></span> | <span data-ttu-id="e27a7-p129">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="e27a7-459">字符串</span><span class="sxs-lookup"><span data-stu-id="e27a7-459">String</span></span> | <span data-ttu-id="e27a7-460">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="e27a7-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="e27a7-461">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="e27a7-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="e27a7-462">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="e27a7-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="e27a7-463">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-463">Requirements</span></span>

|<span data-ttu-id="e27a7-464">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-464">Requirement</span></span>| <span data-ttu-id="e27a7-465">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-466">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-467">1.6</span><span class="sxs-lookup"><span data-stu-id="e27a7-467">1.6</span></span> |
|[<span data-ttu-id="e27a7-468">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-469">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-470">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-471">阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-472">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-472">Example</span></span>

```javascript
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to
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

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="e27a7-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e27a7-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="e27a7-474">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="e27a7-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="e27a7-p131">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-477">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="e27a7-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="e27a7-478">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="e27a7-478">**REST Tokens**</span></span>

<span data-ttu-id="e27a7-p132">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="e27a7-482">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="e27a7-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="e27a7-483">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="e27a7-483">**EWS Tokens**</span></span>

<span data-ttu-id="e27a7-p133">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="e27a7-486">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="e27a7-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-487">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-487">Parameters</span></span>

|<span data-ttu-id="e27a7-488">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-488">Name</span></span>| <span data-ttu-id="e27a7-489">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-489">Type</span></span>| <span data-ttu-id="e27a7-490">属性</span><span class="sxs-lookup"><span data-stu-id="e27a7-490">Attributes</span></span>| <span data-ttu-id="e27a7-491">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="e27a7-492">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-492">Object</span></span> | <span data-ttu-id="e27a7-493">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-493">&lt;optional&gt;</span></span> | <span data-ttu-id="e27a7-494">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e27a7-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="e27a7-495">布尔值</span><span class="sxs-lookup"><span data-stu-id="e27a7-495">Boolean</span></span> |  <span data-ttu-id="e27a7-496">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-496">&lt;optional&gt;</span></span> | <span data-ttu-id="e27a7-p134">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e27a7-499">Object</span><span class="sxs-lookup"><span data-stu-id="e27a7-499">Object</span></span> |  <span data-ttu-id="e27a7-500">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-500">&lt;optional&gt;</span></span> | <span data-ttu-id="e27a7-501">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="e27a7-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="e27a7-502">函数</span><span class="sxs-lookup"><span data-stu-id="e27a7-502">function</span></span>||<span data-ttu-id="e27a7-p135">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-505">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-505">Requirements</span></span>

|<span data-ttu-id="e27a7-506">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-506">Requirement</span></span>| <span data-ttu-id="e27a7-507">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-509">1.5</span><span class="sxs-lookup"><span data-stu-id="e27a7-509">1.5</span></span> |
|[<span data-ttu-id="e27a7-510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-511">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-513">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-514">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-514">Example</span></span>

```javascript
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

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="e27a7-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e27a7-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e27a7-516">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="e27a7-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="e27a7-p136">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="e27a7-p137">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e27a7-522">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="e27a7-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="e27a7-p138">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-525">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-525">Parameters</span></span>

|<span data-ttu-id="e27a7-526">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-526">Name</span></span>| <span data-ttu-id="e27a7-527">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-527">Type</span></span>| <span data-ttu-id="e27a7-528">属性</span><span class="sxs-lookup"><span data-stu-id="e27a7-528">Attributes</span></span>| <span data-ttu-id="e27a7-529">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e27a7-530">函数</span><span class="sxs-lookup"><span data-stu-id="e27a7-530">function</span></span>||<span data-ttu-id="e27a7-p139">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="e27a7-533">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-533">Object</span></span>| <span data-ttu-id="e27a7-534">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-534">&lt;optional&gt;</span></span>|<span data-ttu-id="e27a7-535">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="e27a7-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-536">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-536">Requirements</span></span>

|<span data-ttu-id="e27a7-537">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-537">Requirement</span></span>| <span data-ttu-id="e27a7-538">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-539">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-540">1.3</span><span class="sxs-lookup"><span data-stu-id="e27a7-540">1.3</span></span>|
|[<span data-ttu-id="e27a7-541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-542">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-543">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-544">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-545">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-545">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="e27a7-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e27a7-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e27a7-547">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="e27a7-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="e27a7-548">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="e27a7-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-549">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-549">Parameters</span></span>

|<span data-ttu-id="e27a7-550">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-550">Name</span></span>| <span data-ttu-id="e27a7-551">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-551">Type</span></span>| <span data-ttu-id="e27a7-552">属性</span><span class="sxs-lookup"><span data-stu-id="e27a7-552">Attributes</span></span>| <span data-ttu-id="e27a7-553">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e27a7-554">function</span><span class="sxs-lookup"><span data-stu-id="e27a7-554">function</span></span>||<span data-ttu-id="e27a7-555">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e27a7-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e27a7-556">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="e27a7-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="e27a7-557">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-557">Object</span></span>| <span data-ttu-id="e27a7-558">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-558">&lt;optional&gt;</span></span>|<span data-ttu-id="e27a7-559">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="e27a7-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-560">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-560">Requirements</span></span>

|<span data-ttu-id="e27a7-561">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-561">Requirement</span></span>| <span data-ttu-id="e27a7-562">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-563">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-564">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-564">1.0</span></span>|
|[<span data-ttu-id="e27a7-565">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-566">ReadItem</span></span>|
|[<span data-ttu-id="e27a7-567">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-568">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-569">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-569">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="e27a7-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e27a7-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="e27a7-571">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="e27a7-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-572">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="e27a7-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="e27a7-573">在 iOS 或 Android 上的 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="e27a7-573">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="e27a7-574">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="e27a7-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="e27a7-575">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="e27a7-575">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="e27a7-576">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="e27a7-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="e27a7-577">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="e27a7-577">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="e27a7-578">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="e27a7-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="e27a7-579">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="e27a7-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="e27a7-p141">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="e27a7-582">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="e27a7-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="e27a7-583">版本差异</span><span class="sxs-lookup"><span data-stu-id="e27a7-583">Version differences</span></span>

<span data-ttu-id="e27a7-584">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="e27a7-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="e27a7-p142">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="e27a7-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-588">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-588">Parameters</span></span>

|<span data-ttu-id="e27a7-589">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-589">Name</span></span>| <span data-ttu-id="e27a7-590">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-590">Type</span></span>| <span data-ttu-id="e27a7-591">属性</span><span class="sxs-lookup"><span data-stu-id="e27a7-591">Attributes</span></span>| <span data-ttu-id="e27a7-592">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e27a7-593">字符串</span><span class="sxs-lookup"><span data-stu-id="e27a7-593">String</span></span>||<span data-ttu-id="e27a7-594">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="e27a7-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="e27a7-595">function</span><span class="sxs-lookup"><span data-stu-id="e27a7-595">function</span></span>||<span data-ttu-id="e27a7-596">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e27a7-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e27a7-597">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="e27a7-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="e27a7-598">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="e27a7-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="e27a7-599">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-599">Object</span></span>| <span data-ttu-id="e27a7-600">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-600">&lt;optional&gt;</span></span>|<span data-ttu-id="e27a7-601">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="e27a7-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-602">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-602">Requirements</span></span>

|<span data-ttu-id="e27a7-603">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-603">Requirement</span></span>| <span data-ttu-id="e27a7-604">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-605">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-606">1.0</span><span class="sxs-lookup"><span data-stu-id="e27a7-606">1.0</span></span>|
|[<span data-ttu-id="e27a7-607">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e27a7-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="e27a7-609">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-610">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e27a7-611">示例</span><span class="sxs-lookup"><span data-stu-id="e27a7-611">Example</span></span>

<span data-ttu-id="e27a7-612">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="e27a7-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```javascript
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

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="e27a7-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e27a7-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="e27a7-614">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="e27a7-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="e27a7-615">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="e27a7-615">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e27a7-616">参数</span><span class="sxs-lookup"><span data-stu-id="e27a7-616">Parameters</span></span>

| <span data-ttu-id="e27a7-617">名称</span><span class="sxs-lookup"><span data-stu-id="e27a7-617">Name</span></span> | <span data-ttu-id="e27a7-618">类型</span><span class="sxs-lookup"><span data-stu-id="e27a7-618">Type</span></span> | <span data-ttu-id="e27a7-619">属性</span><span class="sxs-lookup"><span data-stu-id="e27a7-619">Attributes</span></span> | <span data-ttu-id="e27a7-620">说明</span><span class="sxs-lookup"><span data-stu-id="e27a7-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e27a7-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e27a7-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e27a7-622">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="e27a7-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="e27a7-623">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-623">Object</span></span> | <span data-ttu-id="e27a7-624">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-624">&lt;optional&gt;</span></span> | <span data-ttu-id="e27a7-625">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e27a7-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e27a7-626">对象</span><span class="sxs-lookup"><span data-stu-id="e27a7-626">Object</span></span> | <span data-ttu-id="e27a7-627">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-627">&lt;optional&gt;</span></span> | <span data-ttu-id="e27a7-628">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e27a7-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e27a7-629">函数</span><span class="sxs-lookup"><span data-stu-id="e27a7-629">function</span></span>| <span data-ttu-id="e27a7-630">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e27a7-630">&lt;optional&gt;</span></span>|<span data-ttu-id="e27a7-631">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e27a7-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e27a7-632">Requirements</span><span class="sxs-lookup"><span data-stu-id="e27a7-632">Requirements</span></span>

|<span data-ttu-id="e27a7-633">要求</span><span class="sxs-lookup"><span data-stu-id="e27a7-633">Requirement</span></span>| <span data-ttu-id="e27a7-634">值</span><span class="sxs-lookup"><span data-stu-id="e27a7-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="e27a7-635">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e27a7-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e27a7-636">1.5</span><span class="sxs-lookup"><span data-stu-id="e27a7-636">1.5</span></span> |
|[<span data-ttu-id="e27a7-637">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e27a7-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e27a7-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e27a7-638">ReadItem</span></span> |
|[<span data-ttu-id="e27a7-639">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e27a7-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e27a7-640">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e27a7-640">Compose or Read</span></span>|
