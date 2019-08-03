---
title: "\"Context.subname\"-\"邮箱-要求集 1.6\""
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 2ddf26fea8c2285bd577a2f6fb6408431016cc59
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064667"
---
# <a name="mailbox"></a><span data-ttu-id="bc075-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="bc075-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="bc075-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="bc075-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="bc075-104">提供对 Microsoft Outlook 的 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="bc075-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc075-105">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-105">Requirements</span></span>

|<span data-ttu-id="bc075-106">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-106">Requirement</span></span>| <span data-ttu-id="bc075-107">值</span><span class="sxs-lookup"><span data-stu-id="bc075-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-109">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-109">1.0</span></span>|
|[<span data-ttu-id="bc075-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-111">受限</span><span class="sxs-lookup"><span data-stu-id="bc075-111">Restricted</span></span>|
|[<span data-ttu-id="bc075-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bc075-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="bc075-114">Members and methods</span></span>

| <span data-ttu-id="bc075-115">成员</span><span class="sxs-lookup"><span data-stu-id="bc075-115">Member</span></span> | <span data-ttu-id="bc075-116">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bc075-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="bc075-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="bc075-118">成员</span><span class="sxs-lookup"><span data-stu-id="bc075-118">Member</span></span> |
| [<span data-ttu-id="bc075-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="bc075-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="bc075-120">成员</span><span class="sxs-lookup"><span data-stu-id="bc075-120">Member</span></span> |
| [<span data-ttu-id="bc075-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="bc075-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="bc075-122">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-122">Method</span></span> |
| [<span data-ttu-id="bc075-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="bc075-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="bc075-124">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-124">Method</span></span> |
| [<span data-ttu-id="bc075-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="bc075-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="bc075-126">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-126">Method</span></span> |
| [<span data-ttu-id="bc075-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="bc075-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="bc075-128">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-128">Method</span></span> |
| [<span data-ttu-id="bc075-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="bc075-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="bc075-130">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-130">Method</span></span> |
| [<span data-ttu-id="bc075-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="bc075-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="bc075-132">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-132">Method</span></span> |
| [<span data-ttu-id="bc075-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="bc075-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="bc075-134">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-134">Method</span></span> |
| [<span data-ttu-id="bc075-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="bc075-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="bc075-136">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-136">Method</span></span> |
| [<span data-ttu-id="bc075-137">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="bc075-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="bc075-138">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-138">Method</span></span> |
| [<span data-ttu-id="bc075-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="bc075-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="bc075-140">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-140">Method</span></span> |
| [<span data-ttu-id="bc075-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="bc075-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="bc075-142">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-142">Method</span></span> |
| [<span data-ttu-id="bc075-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="bc075-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="bc075-144">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-144">Method</span></span> |
| [<span data-ttu-id="bc075-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="bc075-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="bc075-146">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-146">Method</span></span> |
| [<span data-ttu-id="bc075-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="bc075-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="bc075-148">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="bc075-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="bc075-149">Namespaces</span></span>

<span data-ttu-id="bc075-150">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="bc075-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="bc075-151">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="bc075-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="bc075-152">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="bc075-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="bc075-153">成员</span><span class="sxs-lookup"><span data-stu-id="bc075-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="bc075-154">Mailbox.ewsurl: String</span><span class="sxs-lookup"><span data-stu-id="bc075-154">ewsUrl: String</span></span>

<span data-ttu-id="bc075-155">获取此电子邮件帐户的 Exchange Web Services (EWS) 终点的 URL。</span><span class="sxs-lookup"><span data-stu-id="bc075-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="bc075-156">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="bc075-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-157">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="bc075-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bc075-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="bc075-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="bc075-160">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="bc075-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="bc075-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="bc075-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="bc075-163">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-163">Type</span></span>

*   <span data-ttu-id="bc075-164">String</span><span class="sxs-lookup"><span data-stu-id="bc075-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc075-165">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-165">Requirements</span></span>

|<span data-ttu-id="bc075-166">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-166">Requirement</span></span>| <span data-ttu-id="bc075-167">值</span><span class="sxs-lookup"><span data-stu-id="bc075-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-169">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-169">1.0</span></span>|
|[<span data-ttu-id="bc075-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-171">ReadItem</span></span>|
|[<span data-ttu-id="bc075-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="bc075-174">Office.context.mailbox.resturl: String</span><span class="sxs-lookup"><span data-stu-id="bc075-174">restUrl: String</span></span>

<span data-ttu-id="bc075-175">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="bc075-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="bc075-176">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="bc075-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="bc075-177">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="bc075-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="bc075-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="bc075-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="bc075-180">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-180">Type</span></span>

*   <span data-ttu-id="bc075-181">String</span><span class="sxs-lookup"><span data-stu-id="bc075-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc075-182">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-182">Requirements</span></span>

|<span data-ttu-id="bc075-183">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-183">Requirement</span></span>| <span data-ttu-id="bc075-184">值</span><span class="sxs-lookup"><span data-stu-id="bc075-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-185">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-186">1.5</span><span class="sxs-lookup"><span data-stu-id="bc075-186">1.5</span></span> |
|[<span data-ttu-id="bc075-187">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-188">ReadItem</span></span>|
|[<span data-ttu-id="bc075-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="bc075-191">方法</span><span class="sxs-lookup"><span data-stu-id="bc075-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="bc075-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bc075-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="bc075-193">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="bc075-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="bc075-194">目前，唯一受支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该类型。</span><span class="sxs-lookup"><span data-stu-id="bc075-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="bc075-195">此事件由实现可固定任务窗格的加载项使用，并允许加载项根据当前选定的项目刷新任务窗格 UI。</span><span class="sxs-lookup"><span data-stu-id="bc075-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-196">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-196">Parameters</span></span>

| <span data-ttu-id="bc075-197">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-197">Name</span></span> | <span data-ttu-id="bc075-198">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-198">Type</span></span> | <span data-ttu-id="bc075-199">属性</span><span class="sxs-lookup"><span data-stu-id="bc075-199">Attributes</span></span> | <span data-ttu-id="bc075-200">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="bc075-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="bc075-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="bc075-202">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="bc075-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="bc075-203">函数</span><span class="sxs-lookup"><span data-stu-id="bc075-203">Function</span></span> || <span data-ttu-id="bc075-p106">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="bc075-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="bc075-207">Object</span><span class="sxs-lookup"><span data-stu-id="bc075-207">Object</span></span> | <span data-ttu-id="bc075-208">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-208">&lt;optional&gt;</span></span> | <span data-ttu-id="bc075-209">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="bc075-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="bc075-210">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-210">Object</span></span> | <span data-ttu-id="bc075-211">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-211">&lt;optional&gt;</span></span> | <span data-ttu-id="bc075-212">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="bc075-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="bc075-213">函数</span><span class="sxs-lookup"><span data-stu-id="bc075-213">function</span></span>| <span data-ttu-id="bc075-214">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-214">&lt;optional&gt;</span></span>|<span data-ttu-id="bc075-215">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="bc075-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="bc075-216">Requirements</span></span>

|<span data-ttu-id="bc075-217">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-217">Requirement</span></span>| <span data-ttu-id="bc075-218">值</span><span class="sxs-lookup"><span data-stu-id="bc075-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-219">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-220">1.5</span><span class="sxs-lookup"><span data-stu-id="bc075-220">1.5</span></span> |
|[<span data-ttu-id="bc075-221">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-222">ReadItem</span></span> |
|[<span data-ttu-id="bc075-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-224">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-225">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-225">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="bc075-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="bc075-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="bc075-227">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="bc075-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-228">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="bc075-228">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bc075-p107">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="bc075-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-231">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-231">Parameters</span></span>

|<span data-ttu-id="bc075-232">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-232">Name</span></span>| <span data-ttu-id="bc075-233">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-233">Type</span></span>| <span data-ttu-id="bc075-234">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="bc075-235">字符串</span><span class="sxs-lookup"><span data-stu-id="bc075-235">String</span></span>|<span data-ttu-id="bc075-236">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="bc075-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="bc075-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="bc075-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="bc075-238">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="bc075-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-239">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-239">Requirements</span></span>

|<span data-ttu-id="bc075-240">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-240">Requirement</span></span>| <span data-ttu-id="bc075-241">值</span><span class="sxs-lookup"><span data-stu-id="bc075-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-242">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-243">1.3</span><span class="sxs-lookup"><span data-stu-id="bc075-243">1.3</span></span>|
|[<span data-ttu-id="bc075-244">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-245">受限</span><span class="sxs-lookup"><span data-stu-id="bc075-245">Restricted</span></span>|
|[<span data-ttu-id="bc075-246">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-247">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bc075-248">返回：</span><span class="sxs-lookup"><span data-stu-id="bc075-248">Returns:</span></span>

<span data-ttu-id="bc075-249">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="bc075-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="bc075-250">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="bc075-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="bc075-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="bc075-252">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="bc075-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="bc075-253">适用于桌面或 web 上的 Outlook 的邮件应用程序可以对日期和时间使用不同的时区。</span><span class="sxs-lookup"><span data-stu-id="bc075-253">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="bc075-254">桌面上的 Outlook 使用客户端计算机时区;Web 上的 Outlook 使用 Exchange 管理中心 (EAC) 上设置的时区。</span><span class="sxs-lookup"><span data-stu-id="bc075-254">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="bc075-255">应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="bc075-255">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="bc075-256">如果邮件应用程序在桌面客户端上的 Outlook 中运行, `convertToLocalClientTime`则该方法将返回一个 dictionary 对象, 并将值设置为客户端计算机时区。</span><span class="sxs-lookup"><span data-stu-id="bc075-256">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="bc075-257">如果邮件应用程序在 web 上的 Outlook 中运行, 则`convertToLocalClientTime`该方法将返回一个 dictionary 对象, 其中的值设置为 EAC 中指定的时区。</span><span class="sxs-lookup"><span data-stu-id="bc075-257">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-258">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-258">Parameters</span></span>

|<span data-ttu-id="bc075-259">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-259">Name</span></span>| <span data-ttu-id="bc075-260">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-260">Type</span></span>| <span data-ttu-id="bc075-261">描述</span><span class="sxs-lookup"><span data-stu-id="bc075-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="bc075-262">日期</span><span class="sxs-lookup"><span data-stu-id="bc075-262">Date</span></span>|<span data-ttu-id="bc075-263">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="bc075-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-264">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-264">Requirements</span></span>

|<span data-ttu-id="bc075-265">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-265">Requirement</span></span>| <span data-ttu-id="bc075-266">值</span><span class="sxs-lookup"><span data-stu-id="bc075-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-268">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-268">1.0</span></span>|
|[<span data-ttu-id="bc075-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-270">ReadItem</span></span>|
|[<span data-ttu-id="bc075-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bc075-273">返回：</span><span class="sxs-lookup"><span data-stu-id="bc075-273">Returns:</span></span>

<span data-ttu-id="bc075-274">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="bc075-274">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="bc075-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="bc075-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="bc075-276">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="bc075-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-277">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="bc075-277">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bc075-p110">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="bc075-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-280">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-280">Parameters</span></span>

|<span data-ttu-id="bc075-281">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-281">Name</span></span>| <span data-ttu-id="bc075-282">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-282">Type</span></span>| <span data-ttu-id="bc075-283">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="bc075-284">String</span><span class="sxs-lookup"><span data-stu-id="bc075-284">String</span></span>|<span data-ttu-id="bc075-285">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="bc075-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="bc075-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="bc075-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="bc075-287">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="bc075-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-288">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-288">Requirements</span></span>

|<span data-ttu-id="bc075-289">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-289">Requirement</span></span>| <span data-ttu-id="bc075-290">值</span><span class="sxs-lookup"><span data-stu-id="bc075-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-291">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-292">1.3</span><span class="sxs-lookup"><span data-stu-id="bc075-292">1.3</span></span>|
|[<span data-ttu-id="bc075-293">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-294">受限</span><span class="sxs-lookup"><span data-stu-id="bc075-294">Restricted</span></span>|
|[<span data-ttu-id="bc075-295">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-296">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bc075-297">返回：</span><span class="sxs-lookup"><span data-stu-id="bc075-297">Returns:</span></span>

<span data-ttu-id="bc075-298">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="bc075-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="bc075-299">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="bc075-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="bc075-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="bc075-301">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="bc075-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="bc075-302">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="bc075-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-303">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-303">Parameters</span></span>

|<span data-ttu-id="bc075-304">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-304">Name</span></span>| <span data-ttu-id="bc075-305">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-305">Type</span></span>| <span data-ttu-id="bc075-306">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="bc075-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="bc075-307">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="bc075-308">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="bc075-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-309">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-309">Requirements</span></span>

|<span data-ttu-id="bc075-310">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-310">Requirement</span></span>| <span data-ttu-id="bc075-311">值</span><span class="sxs-lookup"><span data-stu-id="bc075-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-312">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-313">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-313">1.0</span></span>|
|[<span data-ttu-id="bc075-314">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-315">ReadItem</span></span>|
|[<span data-ttu-id="bc075-316">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-317">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bc075-318">返回：</span><span class="sxs-lookup"><span data-stu-id="bc075-318">Returns:</span></span>

<span data-ttu-id="bc075-319">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="bc075-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="bc075-320">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="bc075-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bc075-321">日期</span><span class="sxs-lookup"><span data-stu-id="bc075-321">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="bc075-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="bc075-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="bc075-323">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="bc075-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-324">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="bc075-324">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bc075-325">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="bc075-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="bc075-326">在 Mac 上的 Outlook 中, 可以使用此方法显示不是定期系列的一部分的单个约会, 也可以是定期系列的主约会, 但不能显示该系列的实例。</span><span class="sxs-lookup"><span data-stu-id="bc075-326">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="bc075-327">这是因为在 Mac 上的 Outlook 中, 无法访问定期系列的实例的属性 (包括项目 ID)。</span><span class="sxs-lookup"><span data-stu-id="bc075-327">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="bc075-328">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于32KB 个字符时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="bc075-328">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="bc075-329">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="bc075-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-330">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-330">Parameters</span></span>

|<span data-ttu-id="bc075-331">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-331">Name</span></span>| <span data-ttu-id="bc075-332">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-332">Type</span></span>| <span data-ttu-id="bc075-333">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="bc075-334">字符串</span><span class="sxs-lookup"><span data-stu-id="bc075-334">String</span></span>|<span data-ttu-id="bc075-335">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="bc075-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-336">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-336">Requirements</span></span>

|<span data-ttu-id="bc075-337">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-337">Requirement</span></span>| <span data-ttu-id="bc075-338">值</span><span class="sxs-lookup"><span data-stu-id="bc075-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-339">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-340">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-340">1.0</span></span>|
|[<span data-ttu-id="bc075-341">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-342">ReadItem</span></span>|
|[<span data-ttu-id="bc075-343">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-344">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-345">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="bc075-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="bc075-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="bc075-347">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="bc075-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-348">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="bc075-348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bc075-349">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="bc075-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="bc075-350">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于 32 KB 的字符数时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="bc075-350">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="bc075-351">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="bc075-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="bc075-p112">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="bc075-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-354">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-354">Parameters</span></span>

|<span data-ttu-id="bc075-355">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-355">Name</span></span>| <span data-ttu-id="bc075-356">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-356">Type</span></span>| <span data-ttu-id="bc075-357">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="bc075-358">String</span><span class="sxs-lookup"><span data-stu-id="bc075-358">String</span></span>|<span data-ttu-id="bc075-359">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="bc075-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-360">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-360">Requirements</span></span>

|<span data-ttu-id="bc075-361">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-361">Requirement</span></span>| <span data-ttu-id="bc075-362">值</span><span class="sxs-lookup"><span data-stu-id="bc075-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-363">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-364">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-364">1.0</span></span>|
|[<span data-ttu-id="bc075-365">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-366">ReadItem</span></span>|
|[<span data-ttu-id="bc075-367">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-368">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-369">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="bc075-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="bc075-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="bc075-371">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="bc075-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-372">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="bc075-372">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bc075-p113">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="bc075-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="bc075-375">在 web 和移动设备上的 Outlook 中, 此方法始终显示一个包含 "与会者" 字段的窗体。</span><span class="sxs-lookup"><span data-stu-id="bc075-375">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="bc075-376">如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。</span><span class="sxs-lookup"><span data-stu-id="bc075-376">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="bc075-377">如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="bc075-377">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="bc075-p115">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="bc075-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="bc075-380">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="bc075-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-381">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-382">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="bc075-382">All parameters are optional.</span></span>

|<span data-ttu-id="bc075-383">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-383">Name</span></span>| <span data-ttu-id="bc075-384">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-384">Type</span></span>| <span data-ttu-id="bc075-385">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="bc075-386">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-386">Object</span></span> | <span data-ttu-id="bc075-387">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="bc075-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="bc075-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="bc075-p116">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="bc075-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="bc075-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="bc075-p117">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="bc075-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="bc075-394">Date</span><span class="sxs-lookup"><span data-stu-id="bc075-394">Date</span></span> | <span data-ttu-id="bc075-395">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="bc075-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="bc075-396">Date</span><span class="sxs-lookup"><span data-stu-id="bc075-396">Date</span></span> | <span data-ttu-id="bc075-397">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="bc075-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="bc075-398">字符串</span><span class="sxs-lookup"><span data-stu-id="bc075-398">String</span></span> | <span data-ttu-id="bc075-p118">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="bc075-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="bc075-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="bc075-p119">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="bc075-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="bc075-404">String</span><span class="sxs-lookup"><span data-stu-id="bc075-404">String</span></span> | <span data-ttu-id="bc075-p120">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="bc075-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="bc075-407">字符串</span><span class="sxs-lookup"><span data-stu-id="bc075-407">String</span></span> | <span data-ttu-id="bc075-p121">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="bc075-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bc075-410">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-410">Requirements</span></span>

|<span data-ttu-id="bc075-411">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-411">Requirement</span></span>| <span data-ttu-id="bc075-412">值</span><span class="sxs-lookup"><span data-stu-id="bc075-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-413">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-414">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-414">1.0</span></span>|
|[<span data-ttu-id="bc075-415">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-416">ReadItem</span></span>|
|[<span data-ttu-id="bc075-417">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-418">阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-419">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="bc075-420">Office.context.mailbox.displaynewmessageform (参数)</span><span class="sxs-lookup"><span data-stu-id="bc075-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="bc075-421">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="bc075-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="bc075-422">`displayNewMessageForm`方法将打开一个窗体, 使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="bc075-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="bc075-423">如果指定了参数, 则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="bc075-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="bc075-424">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="bc075-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-425">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-426">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="bc075-426">All parameters are optional.</span></span>

|<span data-ttu-id="bc075-427">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-427">Name</span></span>| <span data-ttu-id="bc075-428">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-428">Type</span></span>| <span data-ttu-id="bc075-429">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="bc075-430">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-430">Object</span></span> | <span data-ttu-id="bc075-431">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="bc075-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="bc075-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="bc075-433">包含电子邮件地址的字符串数组, 或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="bc075-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="bc075-434">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="bc075-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="bc075-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="bc075-436">包含电子邮件地址的字符串数组, 或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="bc075-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="bc075-437">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="bc075-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="bc075-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="bc075-439">包含电子邮件地址的字符串数组, 或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="bc075-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="bc075-440">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="bc075-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="bc075-441">String</span><span class="sxs-lookup"><span data-stu-id="bc075-441">String</span></span> | <span data-ttu-id="bc075-442">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="bc075-442">A string containing the subject of the message.</span></span> <span data-ttu-id="bc075-443">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="bc075-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="bc075-444">String</span><span class="sxs-lookup"><span data-stu-id="bc075-444">String</span></span> | <span data-ttu-id="bc075-445">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="bc075-445">The HTML body of the message.</span></span> <span data-ttu-id="bc075-446">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="bc075-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="bc075-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="bc075-448">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="bc075-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="bc075-449">String</span><span class="sxs-lookup"><span data-stu-id="bc075-449">String</span></span> | <span data-ttu-id="bc075-p128">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="bc075-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="bc075-452">String</span><span class="sxs-lookup"><span data-stu-id="bc075-452">String</span></span> | <span data-ttu-id="bc075-453">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="bc075-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="bc075-454">String</span><span class="sxs-lookup"><span data-stu-id="bc075-454">String</span></span> | <span data-ttu-id="bc075-p129">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="bc075-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="bc075-457">布尔</span><span class="sxs-lookup"><span data-stu-id="bc075-457">Boolean</span></span> | <span data-ttu-id="bc075-p130">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="bc075-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="bc075-460">字符串</span><span class="sxs-lookup"><span data-stu-id="bc075-460">String</span></span> | <span data-ttu-id="bc075-461">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="bc075-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="bc075-462">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="bc075-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="bc075-463">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="bc075-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="bc075-464">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-464">Requirements</span></span>

|<span data-ttu-id="bc075-465">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-465">Requirement</span></span>| <span data-ttu-id="bc075-466">值</span><span class="sxs-lookup"><span data-stu-id="bc075-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-468">1.6</span><span class="sxs-lookup"><span data-stu-id="bc075-468">1.6</span></span> |
|[<span data-ttu-id="bc075-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-470">ReadItem</span></span>|
|[<span data-ttu-id="bc075-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-472">阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-473">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-473">Example</span></span>

```javascript
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="bc075-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="bc075-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="bc075-475">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="bc075-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="bc075-p132">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="bc075-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-478">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="bc075-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="bc075-479">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="bc075-479">**REST Tokens**</span></span>

<span data-ttu-id="bc075-p133">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="bc075-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="bc075-483">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="bc075-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="bc075-484">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="bc075-484">**EWS Tokens**</span></span>

<span data-ttu-id="bc075-p134">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="bc075-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="bc075-487">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="bc075-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-488">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-488">Parameters</span></span>

|<span data-ttu-id="bc075-489">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-489">Name</span></span>| <span data-ttu-id="bc075-490">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-490">Type</span></span>| <span data-ttu-id="bc075-491">属性</span><span class="sxs-lookup"><span data-stu-id="bc075-491">Attributes</span></span>| <span data-ttu-id="bc075-492">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="bc075-493">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-493">Object</span></span> | <span data-ttu-id="bc075-494">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-494">&lt;optional&gt;</span></span> | <span data-ttu-id="bc075-495">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="bc075-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="bc075-496">布尔值</span><span class="sxs-lookup"><span data-stu-id="bc075-496">Boolean</span></span> |  <span data-ttu-id="bc075-497">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-497">&lt;optional&gt;</span></span> | <span data-ttu-id="bc075-p135">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bc075-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="bc075-500">Object</span><span class="sxs-lookup"><span data-stu-id="bc075-500">Object</span></span> |  <span data-ttu-id="bc075-501">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-501">&lt;optional&gt;</span></span> | <span data-ttu-id="bc075-502">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="bc075-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="bc075-503">函数</span><span class="sxs-lookup"><span data-stu-id="bc075-503">function</span></span>||<span data-ttu-id="bc075-p136">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="bc075-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-506">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-506">Requirements</span></span>

|<span data-ttu-id="bc075-507">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-507">Requirement</span></span>| <span data-ttu-id="bc075-508">值</span><span class="sxs-lookup"><span data-stu-id="bc075-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-509">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-510">1.5</span><span class="sxs-lookup"><span data-stu-id="bc075-510">1.5</span></span> |
|[<span data-ttu-id="bc075-511">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-512">ReadItem</span></span>|
|[<span data-ttu-id="bc075-513">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-514">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-515">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="bc075-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bc075-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="bc075-517">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="bc075-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="bc075-p137">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="bc075-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="bc075-p138">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="bc075-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="bc075-523">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="bc075-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="bc075-p139">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="bc075-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-526">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-526">Parameters</span></span>

|<span data-ttu-id="bc075-527">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-527">Name</span></span>| <span data-ttu-id="bc075-528">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-528">Type</span></span>| <span data-ttu-id="bc075-529">属性</span><span class="sxs-lookup"><span data-stu-id="bc075-529">Attributes</span></span>| <span data-ttu-id="bc075-530">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="bc075-531">函数</span><span class="sxs-lookup"><span data-stu-id="bc075-531">function</span></span>||<span data-ttu-id="bc075-p140">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="bc075-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="bc075-534">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-534">Object</span></span>| <span data-ttu-id="bc075-535">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-535">&lt;optional&gt;</span></span>|<span data-ttu-id="bc075-536">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="bc075-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-537">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-537">Requirements</span></span>

|<span data-ttu-id="bc075-538">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-538">Requirement</span></span>| <span data-ttu-id="bc075-539">值</span><span class="sxs-lookup"><span data-stu-id="bc075-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-540">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-541">1.3</span><span class="sxs-lookup"><span data-stu-id="bc075-541">1.3</span></span>|
|[<span data-ttu-id="bc075-542">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-543">ReadItem</span></span>|
|[<span data-ttu-id="bc075-544">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-545">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-546">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-546">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="bc075-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bc075-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="bc075-548">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="bc075-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="bc075-549">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="bc075-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-550">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-550">Parameters</span></span>

|<span data-ttu-id="bc075-551">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-551">Name</span></span>| <span data-ttu-id="bc075-552">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-552">Type</span></span>| <span data-ttu-id="bc075-553">属性</span><span class="sxs-lookup"><span data-stu-id="bc075-553">Attributes</span></span>| <span data-ttu-id="bc075-554">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="bc075-555">function</span><span class="sxs-lookup"><span data-stu-id="bc075-555">function</span></span>||<span data-ttu-id="bc075-556">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="bc075-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bc075-557">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="bc075-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="bc075-558">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-558">Object</span></span>| <span data-ttu-id="bc075-559">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-559">&lt;optional&gt;</span></span>|<span data-ttu-id="bc075-560">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="bc075-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-561">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-561">Requirements</span></span>

|<span data-ttu-id="bc075-562">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-562">Requirement</span></span>| <span data-ttu-id="bc075-563">值</span><span class="sxs-lookup"><span data-stu-id="bc075-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-564">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-565">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-565">1.0</span></span>|
|[<span data-ttu-id="bc075-566">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-567">ReadItem</span></span>|
|[<span data-ttu-id="bc075-568">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-569">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-569">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-570">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-570">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="bc075-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bc075-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="bc075-572">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="bc075-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-573">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="bc075-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="bc075-574">在 iOS 或 Android 上的 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="bc075-574">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="bc075-575">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="bc075-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="bc075-576">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="bc075-576">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="bc075-577">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="bc075-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="bc075-578">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="bc075-578">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="bc075-579">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="bc075-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="bc075-580">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="bc075-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="bc075-p142">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="bc075-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="bc075-583">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="bc075-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="bc075-584">版本差异</span><span class="sxs-lookup"><span data-stu-id="bc075-584">Version differences</span></span>

<span data-ttu-id="bc075-585">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="bc075-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="bc075-p143">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="bc075-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-589">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-589">Parameters</span></span>

|<span data-ttu-id="bc075-590">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-590">Name</span></span>| <span data-ttu-id="bc075-591">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-591">Type</span></span>| <span data-ttu-id="bc075-592">属性</span><span class="sxs-lookup"><span data-stu-id="bc075-592">Attributes</span></span>| <span data-ttu-id="bc075-593">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="bc075-594">字符串</span><span class="sxs-lookup"><span data-stu-id="bc075-594">String</span></span>||<span data-ttu-id="bc075-595">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="bc075-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="bc075-596">function</span><span class="sxs-lookup"><span data-stu-id="bc075-596">function</span></span>||<span data-ttu-id="bc075-597">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="bc075-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bc075-598">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="bc075-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="bc075-599">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="bc075-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="bc075-600">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-600">Object</span></span>| <span data-ttu-id="bc075-601">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-601">&lt;optional&gt;</span></span>|<span data-ttu-id="bc075-602">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="bc075-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-603">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-603">Requirements</span></span>

|<span data-ttu-id="bc075-604">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-604">Requirement</span></span>| <span data-ttu-id="bc075-605">值</span><span class="sxs-lookup"><span data-stu-id="bc075-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-606">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-607">1.0</span><span class="sxs-lookup"><span data-stu-id="bc075-607">1.0</span></span>|
|[<span data-ttu-id="bc075-608">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="bc075-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="bc075-610">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-611">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-611">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc075-612">示例</span><span class="sxs-lookup"><span data-stu-id="bc075-612">Example</span></span>

<span data-ttu-id="bc075-613">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="bc075-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="bc075-614">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bc075-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="bc075-615">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="bc075-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="bc075-616">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="bc075-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bc075-617">参数</span><span class="sxs-lookup"><span data-stu-id="bc075-617">Parameters</span></span>

| <span data-ttu-id="bc075-618">名称</span><span class="sxs-lookup"><span data-stu-id="bc075-618">Name</span></span> | <span data-ttu-id="bc075-619">类型</span><span class="sxs-lookup"><span data-stu-id="bc075-619">Type</span></span> | <span data-ttu-id="bc075-620">属性</span><span class="sxs-lookup"><span data-stu-id="bc075-620">Attributes</span></span> | <span data-ttu-id="bc075-621">说明</span><span class="sxs-lookup"><span data-stu-id="bc075-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="bc075-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="bc075-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="bc075-623">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="bc075-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="bc075-624">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-624">Object</span></span> | <span data-ttu-id="bc075-625">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-625">&lt;optional&gt;</span></span> | <span data-ttu-id="bc075-626">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="bc075-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="bc075-627">对象</span><span class="sxs-lookup"><span data-stu-id="bc075-627">Object</span></span> | <span data-ttu-id="bc075-628">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-628">&lt;optional&gt;</span></span> | <span data-ttu-id="bc075-629">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="bc075-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="bc075-630">函数</span><span class="sxs-lookup"><span data-stu-id="bc075-630">function</span></span>| <span data-ttu-id="bc075-631">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="bc075-631">&lt;optional&gt;</span></span>|<span data-ttu-id="bc075-632">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="bc075-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc075-633">Requirements</span><span class="sxs-lookup"><span data-stu-id="bc075-633">Requirements</span></span>

|<span data-ttu-id="bc075-634">要求</span><span class="sxs-lookup"><span data-stu-id="bc075-634">Requirement</span></span>| <span data-ttu-id="bc075-635">值</span><span class="sxs-lookup"><span data-stu-id="bc075-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc075-636">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="bc075-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc075-637">1.5</span><span class="sxs-lookup"><span data-stu-id="bc075-637">1.5</span></span> |
|[<span data-ttu-id="bc075-638">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="bc075-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc075-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc075-639">ReadItem</span></span> |
|[<span data-ttu-id="bc075-640">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="bc075-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bc075-641">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="bc075-641">Compose or Read</span></span>|
