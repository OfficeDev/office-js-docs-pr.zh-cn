---
title: "\"Context.subname\"-\"邮箱-要求集 1.6\""
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 82a7039602c1896488e6a2358cf345bc157b79de
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695978"
---
# <a name="mailbox"></a><span data-ttu-id="66985-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="66985-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="66985-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="66985-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="66985-104">提供对 Microsoft Outlook 的 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="66985-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="66985-105">要求</span><span class="sxs-lookup"><span data-stu-id="66985-105">Requirements</span></span>

|<span data-ttu-id="66985-106">要求</span><span class="sxs-lookup"><span data-stu-id="66985-106">Requirement</span></span>| <span data-ttu-id="66985-107">值</span><span class="sxs-lookup"><span data-stu-id="66985-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-109">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-109">1.0</span></span>|
|[<span data-ttu-id="66985-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-111">受限</span><span class="sxs-lookup"><span data-stu-id="66985-111">Restricted</span></span>|
|[<span data-ttu-id="66985-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="66985-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="66985-114">Members and methods</span></span>

| <span data-ttu-id="66985-115">成员</span><span class="sxs-lookup"><span data-stu-id="66985-115">Member</span></span> | <span data-ttu-id="66985-116">类型</span><span class="sxs-lookup"><span data-stu-id="66985-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="66985-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="66985-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="66985-118">成员</span><span class="sxs-lookup"><span data-stu-id="66985-118">Member</span></span> |
| [<span data-ttu-id="66985-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="66985-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="66985-120">成员</span><span class="sxs-lookup"><span data-stu-id="66985-120">Member</span></span> |
| [<span data-ttu-id="66985-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="66985-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="66985-122">方法</span><span class="sxs-lookup"><span data-stu-id="66985-122">Method</span></span> |
| [<span data-ttu-id="66985-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="66985-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="66985-124">方法</span><span class="sxs-lookup"><span data-stu-id="66985-124">Method</span></span> |
| [<span data-ttu-id="66985-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="66985-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="66985-126">方法</span><span class="sxs-lookup"><span data-stu-id="66985-126">Method</span></span> |
| [<span data-ttu-id="66985-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="66985-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="66985-128">方法</span><span class="sxs-lookup"><span data-stu-id="66985-128">Method</span></span> |
| [<span data-ttu-id="66985-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="66985-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="66985-130">方法</span><span class="sxs-lookup"><span data-stu-id="66985-130">Method</span></span> |
| [<span data-ttu-id="66985-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="66985-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="66985-132">方法</span><span class="sxs-lookup"><span data-stu-id="66985-132">Method</span></span> |
| [<span data-ttu-id="66985-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="66985-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="66985-134">方法</span><span class="sxs-lookup"><span data-stu-id="66985-134">Method</span></span> |
| [<span data-ttu-id="66985-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="66985-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="66985-136">方法</span><span class="sxs-lookup"><span data-stu-id="66985-136">Method</span></span> |
| [<span data-ttu-id="66985-137">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="66985-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="66985-138">方法</span><span class="sxs-lookup"><span data-stu-id="66985-138">Method</span></span> |
| [<span data-ttu-id="66985-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="66985-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="66985-140">方法</span><span class="sxs-lookup"><span data-stu-id="66985-140">Method</span></span> |
| [<span data-ttu-id="66985-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="66985-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="66985-142">方法</span><span class="sxs-lookup"><span data-stu-id="66985-142">Method</span></span> |
| [<span data-ttu-id="66985-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="66985-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="66985-144">方法</span><span class="sxs-lookup"><span data-stu-id="66985-144">Method</span></span> |
| [<span data-ttu-id="66985-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="66985-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="66985-146">方法</span><span class="sxs-lookup"><span data-stu-id="66985-146">Method</span></span> |
| [<span data-ttu-id="66985-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="66985-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="66985-148">方法</span><span class="sxs-lookup"><span data-stu-id="66985-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="66985-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="66985-149">Namespaces</span></span>

<span data-ttu-id="66985-150">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="66985-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="66985-151">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="66985-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="66985-152">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="66985-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="66985-153">成员</span><span class="sxs-lookup"><span data-stu-id="66985-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="66985-154">Mailbox.ewsurl: String</span><span class="sxs-lookup"><span data-stu-id="66985-154">ewsUrl: String</span></span>

<span data-ttu-id="66985-155">获取此电子邮件帐户的 Exchange Web Services (EWS) 终点的 URL。</span><span class="sxs-lookup"><span data-stu-id="66985-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="66985-156">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="66985-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="66985-157">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="66985-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="66985-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="66985-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="66985-160">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="66985-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="66985-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="66985-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="66985-163">类型</span><span class="sxs-lookup"><span data-stu-id="66985-163">Type</span></span>

*   <span data-ttu-id="66985-164">String</span><span class="sxs-lookup"><span data-stu-id="66985-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="66985-165">要求</span><span class="sxs-lookup"><span data-stu-id="66985-165">Requirements</span></span>

|<span data-ttu-id="66985-166">要求</span><span class="sxs-lookup"><span data-stu-id="66985-166">Requirement</span></span>| <span data-ttu-id="66985-167">值</span><span class="sxs-lookup"><span data-stu-id="66985-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-169">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-169">1.0</span></span>|
|[<span data-ttu-id="66985-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-171">ReadItem</span></span>|
|[<span data-ttu-id="66985-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="66985-174">Office.context.mailbox.resturl: String</span><span class="sxs-lookup"><span data-stu-id="66985-174">restUrl: String</span></span>

<span data-ttu-id="66985-175">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="66985-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="66985-176">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="66985-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="66985-177">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="66985-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="66985-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="66985-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="66985-180">类型</span><span class="sxs-lookup"><span data-stu-id="66985-180">Type</span></span>

*   <span data-ttu-id="66985-181">String</span><span class="sxs-lookup"><span data-stu-id="66985-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="66985-182">要求</span><span class="sxs-lookup"><span data-stu-id="66985-182">Requirements</span></span>

|<span data-ttu-id="66985-183">要求</span><span class="sxs-lookup"><span data-stu-id="66985-183">Requirement</span></span>| <span data-ttu-id="66985-184">值</span><span class="sxs-lookup"><span data-stu-id="66985-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-185">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-186">1.5</span><span class="sxs-lookup"><span data-stu-id="66985-186">1.5</span></span> |
|[<span data-ttu-id="66985-187">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-188">ReadItem</span></span>|
|[<span data-ttu-id="66985-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="66985-191">方法</span><span class="sxs-lookup"><span data-stu-id="66985-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="66985-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="66985-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="66985-193">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="66985-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="66985-194">目前，唯一受支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该类型。</span><span class="sxs-lookup"><span data-stu-id="66985-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="66985-195">此事件由实现可固定任务窗格的加载项使用，并允许加载项根据当前选定的项目刷新任务窗格 UI。</span><span class="sxs-lookup"><span data-stu-id="66985-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-196">参数</span><span class="sxs-lookup"><span data-stu-id="66985-196">Parameters</span></span>

| <span data-ttu-id="66985-197">名称</span><span class="sxs-lookup"><span data-stu-id="66985-197">Name</span></span> | <span data-ttu-id="66985-198">类型</span><span class="sxs-lookup"><span data-stu-id="66985-198">Type</span></span> | <span data-ttu-id="66985-199">属性</span><span class="sxs-lookup"><span data-stu-id="66985-199">Attributes</span></span> | <span data-ttu-id="66985-200">说明</span><span class="sxs-lookup"><span data-stu-id="66985-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="66985-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="66985-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="66985-202">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="66985-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="66985-203">函数</span><span class="sxs-lookup"><span data-stu-id="66985-203">Function</span></span> || <span data-ttu-id="66985-p106">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="66985-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="66985-207">Object</span><span class="sxs-lookup"><span data-stu-id="66985-207">Object</span></span> | <span data-ttu-id="66985-208">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-208">&lt;optional&gt;</span></span> | <span data-ttu-id="66985-209">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="66985-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="66985-210">对象</span><span class="sxs-lookup"><span data-stu-id="66985-210">Object</span></span> | <span data-ttu-id="66985-211">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-211">&lt;optional&gt;</span></span> | <span data-ttu-id="66985-212">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="66985-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="66985-213">函数</span><span class="sxs-lookup"><span data-stu-id="66985-213">function</span></span>| <span data-ttu-id="66985-214">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-214">&lt;optional&gt;</span></span>|<span data-ttu-id="66985-215">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="66985-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="66985-216">Requirements</span></span>

|<span data-ttu-id="66985-217">要求</span><span class="sxs-lookup"><span data-stu-id="66985-217">Requirement</span></span>| <span data-ttu-id="66985-218">值</span><span class="sxs-lookup"><span data-stu-id="66985-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-219">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-220">1.5</span><span class="sxs-lookup"><span data-stu-id="66985-220">1.5</span></span> |
|[<span data-ttu-id="66985-221">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-222">ReadItem</span></span> |
|[<span data-ttu-id="66985-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-224">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-225">示例</span><span class="sxs-lookup"><span data-stu-id="66985-225">Example</span></span>

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
};
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="66985-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="66985-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="66985-227">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="66985-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="66985-228">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="66985-228">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="66985-p107">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="66985-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-231">参数</span><span class="sxs-lookup"><span data-stu-id="66985-231">Parameters</span></span>

|<span data-ttu-id="66985-232">名称</span><span class="sxs-lookup"><span data-stu-id="66985-232">Name</span></span>| <span data-ttu-id="66985-233">类型</span><span class="sxs-lookup"><span data-stu-id="66985-233">Type</span></span>| <span data-ttu-id="66985-234">说明</span><span class="sxs-lookup"><span data-stu-id="66985-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="66985-235">字符串</span><span class="sxs-lookup"><span data-stu-id="66985-235">String</span></span>|<span data-ttu-id="66985-236">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="66985-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="66985-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="66985-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="66985-238">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="66985-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-239">要求</span><span class="sxs-lookup"><span data-stu-id="66985-239">Requirements</span></span>

|<span data-ttu-id="66985-240">要求</span><span class="sxs-lookup"><span data-stu-id="66985-240">Requirement</span></span>| <span data-ttu-id="66985-241">值</span><span class="sxs-lookup"><span data-stu-id="66985-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-242">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-243">1.3</span><span class="sxs-lookup"><span data-stu-id="66985-243">1.3</span></span>|
|[<span data-ttu-id="66985-244">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-245">受限</span><span class="sxs-lookup"><span data-stu-id="66985-245">Restricted</span></span>|
|[<span data-ttu-id="66985-246">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-247">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="66985-248">返回：</span><span class="sxs-lookup"><span data-stu-id="66985-248">Returns:</span></span>

<span data-ttu-id="66985-249">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="66985-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="66985-250">示例</span><span class="sxs-lookup"><span data-stu-id="66985-250">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="66985-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="66985-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="66985-252">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="66985-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="66985-253">适用于桌面或 web 上的 Outlook 的邮件应用程序可以对日期和时间使用不同的时区。</span><span class="sxs-lookup"><span data-stu-id="66985-253">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="66985-254">桌面上的 Outlook 使用客户端计算机时区;Web 上的 Outlook 使用 Exchange 管理中心 (EAC) 上设置的时区。</span><span class="sxs-lookup"><span data-stu-id="66985-254">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="66985-255">应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="66985-255">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="66985-256">如果邮件应用程序在桌面客户端上的 Outlook 中运行, `convertToLocalClientTime`则该方法将返回一个 dictionary 对象, 并将值设置为客户端计算机时区。</span><span class="sxs-lookup"><span data-stu-id="66985-256">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="66985-257">如果邮件应用程序在 web 上的 Outlook 中运行, 则`convertToLocalClientTime`该方法将返回一个 dictionary 对象, 其中的值设置为 EAC 中指定的时区。</span><span class="sxs-lookup"><span data-stu-id="66985-257">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-258">参数</span><span class="sxs-lookup"><span data-stu-id="66985-258">Parameters</span></span>

|<span data-ttu-id="66985-259">名称</span><span class="sxs-lookup"><span data-stu-id="66985-259">Name</span></span>| <span data-ttu-id="66985-260">类型</span><span class="sxs-lookup"><span data-stu-id="66985-260">Type</span></span>| <span data-ttu-id="66985-261">描述</span><span class="sxs-lookup"><span data-stu-id="66985-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="66985-262">日期</span><span class="sxs-lookup"><span data-stu-id="66985-262">Date</span></span>|<span data-ttu-id="66985-263">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="66985-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-264">要求</span><span class="sxs-lookup"><span data-stu-id="66985-264">Requirements</span></span>

|<span data-ttu-id="66985-265">要求</span><span class="sxs-lookup"><span data-stu-id="66985-265">Requirement</span></span>| <span data-ttu-id="66985-266">值</span><span class="sxs-lookup"><span data-stu-id="66985-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-268">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-268">1.0</span></span>|
|[<span data-ttu-id="66985-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-270">ReadItem</span></span>|
|[<span data-ttu-id="66985-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="66985-273">返回：</span><span class="sxs-lookup"><span data-stu-id="66985-273">Returns:</span></span>

<span data-ttu-id="66985-274">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="66985-274">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="66985-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="66985-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="66985-276">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="66985-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="66985-277">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="66985-277">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="66985-p110">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="66985-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-280">参数</span><span class="sxs-lookup"><span data-stu-id="66985-280">Parameters</span></span>

|<span data-ttu-id="66985-281">名称</span><span class="sxs-lookup"><span data-stu-id="66985-281">Name</span></span>| <span data-ttu-id="66985-282">类型</span><span class="sxs-lookup"><span data-stu-id="66985-282">Type</span></span>| <span data-ttu-id="66985-283">说明</span><span class="sxs-lookup"><span data-stu-id="66985-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="66985-284">String</span><span class="sxs-lookup"><span data-stu-id="66985-284">String</span></span>|<span data-ttu-id="66985-285">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="66985-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="66985-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="66985-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="66985-287">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="66985-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-288">要求</span><span class="sxs-lookup"><span data-stu-id="66985-288">Requirements</span></span>

|<span data-ttu-id="66985-289">要求</span><span class="sxs-lookup"><span data-stu-id="66985-289">Requirement</span></span>| <span data-ttu-id="66985-290">值</span><span class="sxs-lookup"><span data-stu-id="66985-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-291">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-292">1.3</span><span class="sxs-lookup"><span data-stu-id="66985-292">1.3</span></span>|
|[<span data-ttu-id="66985-293">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-294">受限</span><span class="sxs-lookup"><span data-stu-id="66985-294">Restricted</span></span>|
|[<span data-ttu-id="66985-295">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-296">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="66985-297">返回：</span><span class="sxs-lookup"><span data-stu-id="66985-297">Returns:</span></span>

<span data-ttu-id="66985-298">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="66985-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="66985-299">示例</span><span class="sxs-lookup"><span data-stu-id="66985-299">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="66985-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="66985-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="66985-301">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="66985-302">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-303">参数</span><span class="sxs-lookup"><span data-stu-id="66985-303">Parameters</span></span>

|<span data-ttu-id="66985-304">名称</span><span class="sxs-lookup"><span data-stu-id="66985-304">Name</span></span>| <span data-ttu-id="66985-305">类型</span><span class="sxs-lookup"><span data-stu-id="66985-305">Type</span></span>| <span data-ttu-id="66985-306">说明</span><span class="sxs-lookup"><span data-stu-id="66985-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="66985-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="66985-307">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="66985-308">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="66985-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-309">要求</span><span class="sxs-lookup"><span data-stu-id="66985-309">Requirements</span></span>

|<span data-ttu-id="66985-310">要求</span><span class="sxs-lookup"><span data-stu-id="66985-310">Requirement</span></span>| <span data-ttu-id="66985-311">值</span><span class="sxs-lookup"><span data-stu-id="66985-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-312">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-313">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-313">1.0</span></span>|
|[<span data-ttu-id="66985-314">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-315">ReadItem</span></span>|
|[<span data-ttu-id="66985-316">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-317">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="66985-318">返回：</span><span class="sxs-lookup"><span data-stu-id="66985-318">Returns:</span></span>

<span data-ttu-id="66985-319">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-319">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="66985-320">类型: Date</span><span class="sxs-lookup"><span data-stu-id="66985-320">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="66985-321">示例</span><span class="sxs-lookup"><span data-stu-id="66985-321">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="66985-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="66985-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="66985-323">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="66985-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="66985-324">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="66985-324">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="66985-325">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="66985-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="66985-326">在 Mac 上的 Outlook 中, 可以使用此方法显示不是定期系列的一部分的单个约会, 也可以是定期系列的主约会, 但不能显示该系列的实例。</span><span class="sxs-lookup"><span data-stu-id="66985-326">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="66985-327">这是因为在 Mac 上的 Outlook 中, 无法访问定期系列的实例的属性 (包括项目 ID)。</span><span class="sxs-lookup"><span data-stu-id="66985-327">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="66985-328">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于32KB 个字符时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="66985-328">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="66985-329">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="66985-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-330">参数</span><span class="sxs-lookup"><span data-stu-id="66985-330">Parameters</span></span>

|<span data-ttu-id="66985-331">名称</span><span class="sxs-lookup"><span data-stu-id="66985-331">Name</span></span>| <span data-ttu-id="66985-332">类型</span><span class="sxs-lookup"><span data-stu-id="66985-332">Type</span></span>| <span data-ttu-id="66985-333">说明</span><span class="sxs-lookup"><span data-stu-id="66985-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="66985-334">字符串</span><span class="sxs-lookup"><span data-stu-id="66985-334">String</span></span>|<span data-ttu-id="66985-335">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="66985-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-336">要求</span><span class="sxs-lookup"><span data-stu-id="66985-336">Requirements</span></span>

|<span data-ttu-id="66985-337">要求</span><span class="sxs-lookup"><span data-stu-id="66985-337">Requirement</span></span>| <span data-ttu-id="66985-338">值</span><span class="sxs-lookup"><span data-stu-id="66985-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-339">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-340">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-340">1.0</span></span>|
|[<span data-ttu-id="66985-341">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-342">ReadItem</span></span>|
|[<span data-ttu-id="66985-343">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-344">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-345">示例</span><span class="sxs-lookup"><span data-stu-id="66985-345">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="66985-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="66985-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="66985-347">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="66985-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="66985-348">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="66985-348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="66985-349">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="66985-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="66985-350">在 web 上的 Outlook 中, 仅当窗体的正文小于或等于 32 KB 的字符数时, 此方法才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="66985-350">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="66985-351">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="66985-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="66985-p112">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="66985-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-354">参数</span><span class="sxs-lookup"><span data-stu-id="66985-354">Parameters</span></span>

|<span data-ttu-id="66985-355">名称</span><span class="sxs-lookup"><span data-stu-id="66985-355">Name</span></span>| <span data-ttu-id="66985-356">类型</span><span class="sxs-lookup"><span data-stu-id="66985-356">Type</span></span>| <span data-ttu-id="66985-357">说明</span><span class="sxs-lookup"><span data-stu-id="66985-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="66985-358">String</span><span class="sxs-lookup"><span data-stu-id="66985-358">String</span></span>|<span data-ttu-id="66985-359">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="66985-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-360">要求</span><span class="sxs-lookup"><span data-stu-id="66985-360">Requirements</span></span>

|<span data-ttu-id="66985-361">要求</span><span class="sxs-lookup"><span data-stu-id="66985-361">Requirement</span></span>| <span data-ttu-id="66985-362">值</span><span class="sxs-lookup"><span data-stu-id="66985-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-363">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-364">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-364">1.0</span></span>|
|[<span data-ttu-id="66985-365">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-366">ReadItem</span></span>|
|[<span data-ttu-id="66985-367">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-368">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-369">示例</span><span class="sxs-lookup"><span data-stu-id="66985-369">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="66985-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="66985-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="66985-371">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="66985-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="66985-372">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="66985-372">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="66985-p113">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="66985-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="66985-375">在 web 和移动设备上的 Outlook 中, 此方法始终显示一个包含 "与会者" 字段的窗体。</span><span class="sxs-lookup"><span data-stu-id="66985-375">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="66985-376">如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。</span><span class="sxs-lookup"><span data-stu-id="66985-376">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="66985-377">如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="66985-377">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="66985-p115">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="66985-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="66985-380">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="66985-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-381">参数</span><span class="sxs-lookup"><span data-stu-id="66985-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="66985-382">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="66985-382">All parameters are optional.</span></span>

|<span data-ttu-id="66985-383">名称</span><span class="sxs-lookup"><span data-stu-id="66985-383">Name</span></span>| <span data-ttu-id="66985-384">类型</span><span class="sxs-lookup"><span data-stu-id="66985-384">Type</span></span>| <span data-ttu-id="66985-385">说明</span><span class="sxs-lookup"><span data-stu-id="66985-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="66985-386">对象</span><span class="sxs-lookup"><span data-stu-id="66985-386">Object</span></span> | <span data-ttu-id="66985-387">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="66985-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="66985-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="66985-p116">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="66985-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="66985-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="66985-p117">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="66985-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="66985-394">Date</span><span class="sxs-lookup"><span data-stu-id="66985-394">Date</span></span> | <span data-ttu-id="66985-395">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="66985-396">Date</span><span class="sxs-lookup"><span data-stu-id="66985-396">Date</span></span> | <span data-ttu-id="66985-397">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="66985-398">字符串</span><span class="sxs-lookup"><span data-stu-id="66985-398">String</span></span> | <span data-ttu-id="66985-p118">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="66985-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="66985-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="66985-p119">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="66985-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="66985-404">String</span><span class="sxs-lookup"><span data-stu-id="66985-404">String</span></span> | <span data-ttu-id="66985-p120">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="66985-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="66985-407">字符串</span><span class="sxs-lookup"><span data-stu-id="66985-407">String</span></span> | <span data-ttu-id="66985-p121">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="66985-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="66985-410">要求</span><span class="sxs-lookup"><span data-stu-id="66985-410">Requirements</span></span>

|<span data-ttu-id="66985-411">要求</span><span class="sxs-lookup"><span data-stu-id="66985-411">Requirement</span></span>| <span data-ttu-id="66985-412">值</span><span class="sxs-lookup"><span data-stu-id="66985-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-413">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-414">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-414">1.0</span></span>|
|[<span data-ttu-id="66985-415">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-416">ReadItem</span></span>|
|[<span data-ttu-id="66985-417">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-418">阅读</span><span class="sxs-lookup"><span data-stu-id="66985-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-419">示例</span><span class="sxs-lookup"><span data-stu-id="66985-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="66985-420">Office.context.mailbox.displaynewmessageform (参数)</span><span class="sxs-lookup"><span data-stu-id="66985-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="66985-421">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="66985-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="66985-422">`displayNewMessageForm`方法将打开一个窗体, 使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="66985-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="66985-423">如果指定了参数, 则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="66985-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="66985-424">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="66985-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-425">参数</span><span class="sxs-lookup"><span data-stu-id="66985-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="66985-426">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="66985-426">All parameters are optional.</span></span>

|<span data-ttu-id="66985-427">名称</span><span class="sxs-lookup"><span data-stu-id="66985-427">Name</span></span>| <span data-ttu-id="66985-428">类型</span><span class="sxs-lookup"><span data-stu-id="66985-428">Type</span></span>| <span data-ttu-id="66985-429">说明</span><span class="sxs-lookup"><span data-stu-id="66985-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="66985-430">对象</span><span class="sxs-lookup"><span data-stu-id="66985-430">Object</span></span> | <span data-ttu-id="66985-431">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="66985-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="66985-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="66985-433">包含电子邮件地址的字符串数组, 或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="66985-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="66985-434">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="66985-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="66985-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="66985-436">包含电子邮件地址的字符串数组, 或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="66985-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="66985-437">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="66985-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="66985-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="66985-439">包含电子邮件地址的字符串数组, 或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="66985-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="66985-440">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="66985-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="66985-441">String</span><span class="sxs-lookup"><span data-stu-id="66985-441">String</span></span> | <span data-ttu-id="66985-442">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="66985-442">A string containing the subject of the message.</span></span> <span data-ttu-id="66985-443">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="66985-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="66985-444">String</span><span class="sxs-lookup"><span data-stu-id="66985-444">String</span></span> | <span data-ttu-id="66985-445">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="66985-445">The HTML body of the message.</span></span> <span data-ttu-id="66985-446">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="66985-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="66985-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="66985-448">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="66985-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="66985-449">String</span><span class="sxs-lookup"><span data-stu-id="66985-449">String</span></span> | <span data-ttu-id="66985-p128">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="66985-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="66985-452">String</span><span class="sxs-lookup"><span data-stu-id="66985-452">String</span></span> | <span data-ttu-id="66985-453">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="66985-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="66985-454">String</span><span class="sxs-lookup"><span data-stu-id="66985-454">String</span></span> | <span data-ttu-id="66985-p129">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="66985-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="66985-457">布尔</span><span class="sxs-lookup"><span data-stu-id="66985-457">Boolean</span></span> | <span data-ttu-id="66985-p130">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="66985-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="66985-460">字符串</span><span class="sxs-lookup"><span data-stu-id="66985-460">String</span></span> | <span data-ttu-id="66985-461">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="66985-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="66985-462">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="66985-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="66985-463">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="66985-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="66985-464">要求</span><span class="sxs-lookup"><span data-stu-id="66985-464">Requirements</span></span>

|<span data-ttu-id="66985-465">要求</span><span class="sxs-lookup"><span data-stu-id="66985-465">Requirement</span></span>| <span data-ttu-id="66985-466">值</span><span class="sxs-lookup"><span data-stu-id="66985-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-468">1.6</span><span class="sxs-lookup"><span data-stu-id="66985-468">1.6</span></span> |
|[<span data-ttu-id="66985-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-470">ReadItem</span></span>|
|[<span data-ttu-id="66985-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-472">阅读</span><span class="sxs-lookup"><span data-stu-id="66985-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-473">示例</span><span class="sxs-lookup"><span data-stu-id="66985-473">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="66985-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="66985-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="66985-475">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="66985-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="66985-p132">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="66985-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="66985-478">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="66985-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="66985-479">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="66985-479">**REST Tokens**</span></span>

<span data-ttu-id="66985-p133">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="66985-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="66985-483">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="66985-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="66985-484">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="66985-484">**EWS Tokens**</span></span>

<span data-ttu-id="66985-p134">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="66985-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="66985-487">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="66985-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-488">参数</span><span class="sxs-lookup"><span data-stu-id="66985-488">Parameters</span></span>

|<span data-ttu-id="66985-489">名称</span><span class="sxs-lookup"><span data-stu-id="66985-489">Name</span></span>| <span data-ttu-id="66985-490">类型</span><span class="sxs-lookup"><span data-stu-id="66985-490">Type</span></span>| <span data-ttu-id="66985-491">属性</span><span class="sxs-lookup"><span data-stu-id="66985-491">Attributes</span></span>| <span data-ttu-id="66985-492">说明</span><span class="sxs-lookup"><span data-stu-id="66985-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="66985-493">对象</span><span class="sxs-lookup"><span data-stu-id="66985-493">Object</span></span> | <span data-ttu-id="66985-494">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-494">&lt;optional&gt;</span></span> | <span data-ttu-id="66985-495">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="66985-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="66985-496">布尔值</span><span class="sxs-lookup"><span data-stu-id="66985-496">Boolean</span></span> |  <span data-ttu-id="66985-497">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-497">&lt;optional&gt;</span></span> | <span data-ttu-id="66985-p135">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="66985-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="66985-500">Object</span><span class="sxs-lookup"><span data-stu-id="66985-500">Object</span></span> |  <span data-ttu-id="66985-501">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-501">&lt;optional&gt;</span></span> | <span data-ttu-id="66985-502">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="66985-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="66985-503">函数</span><span class="sxs-lookup"><span data-stu-id="66985-503">function</span></span>||<span data-ttu-id="66985-504">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="66985-504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="66985-505">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="66985-505">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="66985-506">如果出现错误, 则`asyncResult.error`和`asyncResult.diagnostics`属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="66985-506">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="66985-507">错误</span><span class="sxs-lookup"><span data-stu-id="66985-507">Errors</span></span>

|<span data-ttu-id="66985-508">错误代码</span><span class="sxs-lookup"><span data-stu-id="66985-508">Error code</span></span>|<span data-ttu-id="66985-509">说明</span><span class="sxs-lookup"><span data-stu-id="66985-509">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="66985-510">请求失败。</span><span class="sxs-lookup"><span data-stu-id="66985-510">The request has failed.</span></span> <span data-ttu-id="66985-511">请查看 HTTP 错误代码的 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-511">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="66985-512">Exchange 服务器返回错误。</span><span class="sxs-lookup"><span data-stu-id="66985-512">The Exchange server returned an error.</span></span> <span data-ttu-id="66985-513">有关详细信息, 请参阅 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-513">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="66985-514">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="66985-514">The user is no longer connected to the network.</span></span> <span data-ttu-id="66985-515">请检查你的网络连接, 然后重试。</span><span class="sxs-lookup"><span data-stu-id="66985-515">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-516">要求</span><span class="sxs-lookup"><span data-stu-id="66985-516">Requirements</span></span>

|<span data-ttu-id="66985-517">要求</span><span class="sxs-lookup"><span data-stu-id="66985-517">Requirement</span></span>| <span data-ttu-id="66985-518">值</span><span class="sxs-lookup"><span data-stu-id="66985-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-519">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-520">1.5</span><span class="sxs-lookup"><span data-stu-id="66985-520">1.5</span></span> |
|[<span data-ttu-id="66985-521">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-522">ReadItem</span></span>|
|[<span data-ttu-id="66985-523">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-524">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="66985-524">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-525">示例</span><span class="sxs-lookup"><span data-stu-id="66985-525">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="66985-526">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="66985-526">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="66985-527">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="66985-527">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="66985-p139">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="66985-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="66985-p140">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="66985-p140">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="66985-533">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="66985-533">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="66985-p141">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="66985-p141">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-536">参数</span><span class="sxs-lookup"><span data-stu-id="66985-536">Parameters</span></span>

|<span data-ttu-id="66985-537">名称</span><span class="sxs-lookup"><span data-stu-id="66985-537">Name</span></span>| <span data-ttu-id="66985-538">类型</span><span class="sxs-lookup"><span data-stu-id="66985-538">Type</span></span>| <span data-ttu-id="66985-539">属性</span><span class="sxs-lookup"><span data-stu-id="66985-539">Attributes</span></span>| <span data-ttu-id="66985-540">说明</span><span class="sxs-lookup"><span data-stu-id="66985-540">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="66985-541">function</span><span class="sxs-lookup"><span data-stu-id="66985-541">function</span></span>||<span data-ttu-id="66985-542">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="66985-542">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="66985-543">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="66985-543">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="66985-544">如果出现错误, 则`asyncResult.error`和`asyncResult.diagnostics`属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="66985-544">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="66985-545">对象</span><span class="sxs-lookup"><span data-stu-id="66985-545">Object</span></span>| <span data-ttu-id="66985-546">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-546">&lt;optional&gt;</span></span>|<span data-ttu-id="66985-547">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="66985-547">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="66985-548">错误</span><span class="sxs-lookup"><span data-stu-id="66985-548">Errors</span></span>

|<span data-ttu-id="66985-549">错误代码</span><span class="sxs-lookup"><span data-stu-id="66985-549">Error code</span></span>|<span data-ttu-id="66985-550">说明</span><span class="sxs-lookup"><span data-stu-id="66985-550">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="66985-551">请求失败。</span><span class="sxs-lookup"><span data-stu-id="66985-551">The request has failed.</span></span> <span data-ttu-id="66985-552">请查看 HTTP 错误代码的 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-552">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="66985-553">Exchange 服务器返回错误。</span><span class="sxs-lookup"><span data-stu-id="66985-553">The Exchange server returned an error.</span></span> <span data-ttu-id="66985-554">有关详细信息, 请参阅 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-554">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="66985-555">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="66985-555">The user is no longer connected to the network.</span></span> <span data-ttu-id="66985-556">请检查你的网络连接, 然后重试。</span><span class="sxs-lookup"><span data-stu-id="66985-556">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-557">要求</span><span class="sxs-lookup"><span data-stu-id="66985-557">Requirements</span></span>

|<span data-ttu-id="66985-558">要求</span><span class="sxs-lookup"><span data-stu-id="66985-558">Requirement</span></span>| <span data-ttu-id="66985-559">值</span><span class="sxs-lookup"><span data-stu-id="66985-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-560">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-561">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-561">1.0</span></span>|
|[<span data-ttu-id="66985-562">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-562">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-563">ReadItem</span></span>|
|[<span data-ttu-id="66985-564">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-564">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-565">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="66985-565">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-566">示例</span><span class="sxs-lookup"><span data-stu-id="66985-566">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="66985-567">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="66985-567">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="66985-568">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="66985-568">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="66985-569">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="66985-569">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-570">参数</span><span class="sxs-lookup"><span data-stu-id="66985-570">Parameters</span></span>

|<span data-ttu-id="66985-571">名称</span><span class="sxs-lookup"><span data-stu-id="66985-571">Name</span></span>| <span data-ttu-id="66985-572">类型</span><span class="sxs-lookup"><span data-stu-id="66985-572">Type</span></span>| <span data-ttu-id="66985-573">属性</span><span class="sxs-lookup"><span data-stu-id="66985-573">Attributes</span></span>| <span data-ttu-id="66985-574">说明</span><span class="sxs-lookup"><span data-stu-id="66985-574">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="66985-575">function</span><span class="sxs-lookup"><span data-stu-id="66985-575">function</span></span>||<span data-ttu-id="66985-576">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="66985-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="66985-577">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="66985-577">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="66985-578">如果出现错误, 则`asyncResult.error`和`asyncResult.diagnostics`属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="66985-578">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="66985-579">对象</span><span class="sxs-lookup"><span data-stu-id="66985-579">Object</span></span>| <span data-ttu-id="66985-580">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-580">&lt;optional&gt;</span></span>|<span data-ttu-id="66985-581">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="66985-581">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="66985-582">错误</span><span class="sxs-lookup"><span data-stu-id="66985-582">Errors</span></span>

|<span data-ttu-id="66985-583">错误代码</span><span class="sxs-lookup"><span data-stu-id="66985-583">Error code</span></span>|<span data-ttu-id="66985-584">说明</span><span class="sxs-lookup"><span data-stu-id="66985-584">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="66985-585">请求失败。</span><span class="sxs-lookup"><span data-stu-id="66985-585">The request has failed.</span></span> <span data-ttu-id="66985-586">请查看 HTTP 错误代码的 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-586">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="66985-587">Exchange 服务器返回错误。</span><span class="sxs-lookup"><span data-stu-id="66985-587">The Exchange server returned an error.</span></span> <span data-ttu-id="66985-588">有关详细信息, 请参阅 diagnostics 对象。</span><span class="sxs-lookup"><span data-stu-id="66985-588">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="66985-589">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="66985-589">The user is no longer connected to the network.</span></span> <span data-ttu-id="66985-590">请检查你的网络连接, 然后重试。</span><span class="sxs-lookup"><span data-stu-id="66985-590">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-591">要求</span><span class="sxs-lookup"><span data-stu-id="66985-591">Requirements</span></span>

|<span data-ttu-id="66985-592">要求</span><span class="sxs-lookup"><span data-stu-id="66985-592">Requirement</span></span>| <span data-ttu-id="66985-593">值</span><span class="sxs-lookup"><span data-stu-id="66985-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-594">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-595">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-595">1.0</span></span>|
|[<span data-ttu-id="66985-596">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-596">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-597">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-597">ReadItem</span></span>|
|[<span data-ttu-id="66985-598">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-598">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-599">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-599">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-600">示例</span><span class="sxs-lookup"><span data-stu-id="66985-600">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="66985-601">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="66985-601">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="66985-602">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="66985-602">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="66985-603">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="66985-603">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="66985-604">在 iOS 或 Android 上的 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="66985-604">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="66985-605">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="66985-605">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="66985-606">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="66985-606">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="66985-607">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="66985-607">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="66985-608">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="66985-608">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="66985-609">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="66985-609">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="66985-610">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="66985-610">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="66985-p149">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="66985-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="66985-613">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="66985-613">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="66985-614">版本差异</span><span class="sxs-lookup"><span data-stu-id="66985-614">Version differences</span></span>

<span data-ttu-id="66985-615">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="66985-615">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="66985-p150">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="66985-p150">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-619">参数</span><span class="sxs-lookup"><span data-stu-id="66985-619">Parameters</span></span>

|<span data-ttu-id="66985-620">名称</span><span class="sxs-lookup"><span data-stu-id="66985-620">Name</span></span>| <span data-ttu-id="66985-621">类型</span><span class="sxs-lookup"><span data-stu-id="66985-621">Type</span></span>| <span data-ttu-id="66985-622">属性</span><span class="sxs-lookup"><span data-stu-id="66985-622">Attributes</span></span>| <span data-ttu-id="66985-623">说明</span><span class="sxs-lookup"><span data-stu-id="66985-623">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="66985-624">字符串</span><span class="sxs-lookup"><span data-stu-id="66985-624">String</span></span>||<span data-ttu-id="66985-625">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="66985-625">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="66985-626">函数</span><span class="sxs-lookup"><span data-stu-id="66985-626">function</span></span>||<span data-ttu-id="66985-627">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="66985-627">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="66985-628">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="66985-628">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="66985-629">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="66985-629">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="66985-630">对象</span><span class="sxs-lookup"><span data-stu-id="66985-630">Object</span></span>| <span data-ttu-id="66985-631">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-631">&lt;optional&gt;</span></span>|<span data-ttu-id="66985-632">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="66985-632">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-633">要求</span><span class="sxs-lookup"><span data-stu-id="66985-633">Requirements</span></span>

|<span data-ttu-id="66985-634">要求</span><span class="sxs-lookup"><span data-stu-id="66985-634">Requirement</span></span>| <span data-ttu-id="66985-635">值</span><span class="sxs-lookup"><span data-stu-id="66985-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-636">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-637">1.0</span><span class="sxs-lookup"><span data-stu-id="66985-637">1.0</span></span>|
|[<span data-ttu-id="66985-638">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-639">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="66985-639">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="66985-640">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-641">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-641">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="66985-642">示例</span><span class="sxs-lookup"><span data-stu-id="66985-642">Example</span></span>

<span data-ttu-id="66985-643">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="66985-643">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="66985-644">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="66985-644">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="66985-645">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="66985-645">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="66985-646">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="66985-646">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="66985-647">参数</span><span class="sxs-lookup"><span data-stu-id="66985-647">Parameters</span></span>

| <span data-ttu-id="66985-648">名称</span><span class="sxs-lookup"><span data-stu-id="66985-648">Name</span></span> | <span data-ttu-id="66985-649">类型</span><span class="sxs-lookup"><span data-stu-id="66985-649">Type</span></span> | <span data-ttu-id="66985-650">属性</span><span class="sxs-lookup"><span data-stu-id="66985-650">Attributes</span></span> | <span data-ttu-id="66985-651">说明</span><span class="sxs-lookup"><span data-stu-id="66985-651">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="66985-652">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="66985-652">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="66985-653">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="66985-653">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="66985-654">对象</span><span class="sxs-lookup"><span data-stu-id="66985-654">Object</span></span> | <span data-ttu-id="66985-655">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-655">&lt;optional&gt;</span></span> | <span data-ttu-id="66985-656">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="66985-656">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="66985-657">对象</span><span class="sxs-lookup"><span data-stu-id="66985-657">Object</span></span> | <span data-ttu-id="66985-658">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-658">&lt;optional&gt;</span></span> | <span data-ttu-id="66985-659">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="66985-659">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="66985-660">函数</span><span class="sxs-lookup"><span data-stu-id="66985-660">function</span></span>| <span data-ttu-id="66985-661">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="66985-661">&lt;optional&gt;</span></span>|<span data-ttu-id="66985-662">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="66985-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="66985-663">Requirements</span><span class="sxs-lookup"><span data-stu-id="66985-663">Requirements</span></span>

|<span data-ttu-id="66985-664">要求</span><span class="sxs-lookup"><span data-stu-id="66985-664">Requirement</span></span>| <span data-ttu-id="66985-665">值</span><span class="sxs-lookup"><span data-stu-id="66985-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="66985-666">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="66985-666">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66985-667">1.5</span><span class="sxs-lookup"><span data-stu-id="66985-667">1.5</span></span> |
|[<span data-ttu-id="66985-668">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="66985-668">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66985-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66985-669">ReadItem</span></span> |
|[<span data-ttu-id="66985-670">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="66985-670">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="66985-671">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="66985-671">Compose or Read</span></span>|
