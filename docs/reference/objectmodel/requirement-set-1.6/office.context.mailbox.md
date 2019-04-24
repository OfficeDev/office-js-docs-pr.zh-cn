---
title: "\"context.subname\"-\"邮箱-要求集 1.6\""
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 9b91a61d301434886723a55eca9608f004f598eb
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451778"
---
# <a name="mailbox"></a><span data-ttu-id="359f6-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="359f6-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="359f6-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="359f6-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="359f6-104">为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="359f6-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="359f6-105">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-105">Requirements</span></span>

|<span data-ttu-id="359f6-106">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-106">Requirement</span></span>| <span data-ttu-id="359f6-107">值</span><span class="sxs-lookup"><span data-stu-id="359f6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-109">1.0</span></span>|
|[<span data-ttu-id="359f6-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-111">受限</span><span class="sxs-lookup"><span data-stu-id="359f6-111">Restricted</span></span>|
|[<span data-ttu-id="359f6-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="359f6-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="359f6-114">Members and methods</span></span>

| <span data-ttu-id="359f6-115">成员</span><span class="sxs-lookup"><span data-stu-id="359f6-115">Member</span></span> | <span data-ttu-id="359f6-116">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="359f6-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="359f6-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="359f6-118">成员</span><span class="sxs-lookup"><span data-stu-id="359f6-118">Member</span></span> |
| [<span data-ttu-id="359f6-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="359f6-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="359f6-120">成员</span><span class="sxs-lookup"><span data-stu-id="359f6-120">Member</span></span> |
| [<span data-ttu-id="359f6-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="359f6-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="359f6-122">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-122">Method</span></span> |
| [<span data-ttu-id="359f6-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="359f6-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="359f6-124">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-124">Method</span></span> |
| [<span data-ttu-id="359f6-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="359f6-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="359f6-126">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-126">Method</span></span> |
| [<span data-ttu-id="359f6-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="359f6-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="359f6-128">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-128">Method</span></span> |
| [<span data-ttu-id="359f6-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="359f6-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="359f6-130">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-130">Method</span></span> |
| [<span data-ttu-id="359f6-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="359f6-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="359f6-132">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-132">Method</span></span> |
| [<span data-ttu-id="359f6-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="359f6-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="359f6-134">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-134">Method</span></span> |
| [<span data-ttu-id="359f6-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="359f6-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="359f6-136">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-136">Method</span></span> |
| [<span data-ttu-id="359f6-137">office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="359f6-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="359f6-138">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-138">Method</span></span> |
| [<span data-ttu-id="359f6-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="359f6-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="359f6-140">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-140">Method</span></span> |
| [<span data-ttu-id="359f6-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="359f6-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="359f6-142">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-142">Method</span></span> |
| [<span data-ttu-id="359f6-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="359f6-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="359f6-144">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-144">Method</span></span> |
| [<span data-ttu-id="359f6-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="359f6-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="359f6-146">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-146">Method</span></span> |
| [<span data-ttu-id="359f6-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="359f6-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="359f6-148">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="359f6-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="359f6-149">Namespaces</span></span>

<span data-ttu-id="359f6-150">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="359f6-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="359f6-151">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="359f6-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="359f6-152">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="359f6-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="359f6-153">成员</span><span class="sxs-lookup"><span data-stu-id="359f6-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="359f6-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="359f6-154">ewsUrl :String</span></span>

<span data-ttu-id="359f6-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="359f6-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-157">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="359f6-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="359f6-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="359f6-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="359f6-160">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="359f6-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="359f6-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="359f6-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="359f6-163">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-163">Type</span></span>

*   <span data-ttu-id="359f6-164">String</span><span class="sxs-lookup"><span data-stu-id="359f6-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="359f6-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-165">Requirements</span></span>

|<span data-ttu-id="359f6-166">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-166">Requirement</span></span>| <span data-ttu-id="359f6-167">值</span><span class="sxs-lookup"><span data-stu-id="359f6-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-169">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-169">1.0</span></span>|
|[<span data-ttu-id="359f6-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-171">ReadItem</span></span>|
|[<span data-ttu-id="359f6-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="359f6-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="359f6-174">restUrl :String</span></span>

<span data-ttu-id="359f6-175">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="359f6-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="359f6-176">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="359f6-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="359f6-177">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="359f6-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="359f6-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="359f6-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="359f6-180">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-180">Type</span></span>

*   <span data-ttu-id="359f6-181">String</span><span class="sxs-lookup"><span data-stu-id="359f6-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="359f6-182">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-182">Requirements</span></span>

|<span data-ttu-id="359f6-183">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-183">Requirement</span></span>| <span data-ttu-id="359f6-184">值</span><span class="sxs-lookup"><span data-stu-id="359f6-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-185">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-186">1.5</span><span class="sxs-lookup"><span data-stu-id="359f6-186">1.5</span></span> |
|[<span data-ttu-id="359f6-187">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-188">ReadItem</span></span>|
|[<span data-ttu-id="359f6-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="359f6-191">方法</span><span class="sxs-lookup"><span data-stu-id="359f6-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="359f6-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="359f6-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="359f6-193">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="359f6-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="359f6-194">目前，唯一受支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该类型。</span><span class="sxs-lookup"><span data-stu-id="359f6-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="359f6-195">此事件由实现可固定任务窗格的加载项使用，并允许加载项根据当前选定的项目刷新任务窗格 UI。</span><span class="sxs-lookup"><span data-stu-id="359f6-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-196">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-196">Parameters</span></span>

| <span data-ttu-id="359f6-197">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-197">Name</span></span> | <span data-ttu-id="359f6-198">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-198">Type</span></span> | <span data-ttu-id="359f6-199">属性</span><span class="sxs-lookup"><span data-stu-id="359f6-199">Attributes</span></span> | <span data-ttu-id="359f6-200">说明</span><span class="sxs-lookup"><span data-stu-id="359f6-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="359f6-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="359f6-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="359f6-202">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="359f6-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="359f6-203">函数</span><span class="sxs-lookup"><span data-stu-id="359f6-203">Function</span></span> || <span data-ttu-id="359f6-p106">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="359f6-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="359f6-207">Object</span><span class="sxs-lookup"><span data-stu-id="359f6-207">Object</span></span> | <span data-ttu-id="359f6-208">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-208">&lt;optional&gt;</span></span> | <span data-ttu-id="359f6-209">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="359f6-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="359f6-210">对象</span><span class="sxs-lookup"><span data-stu-id="359f6-210">Object</span></span> | <span data-ttu-id="359f6-211">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-211">&lt;optional&gt;</span></span> | <span data-ttu-id="359f6-212">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="359f6-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="359f6-213">函数</span><span class="sxs-lookup"><span data-stu-id="359f6-213">function</span></span>| <span data-ttu-id="359f6-214">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-214">&lt;optional&gt;</span></span>|<span data-ttu-id="359f6-215">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="359f6-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-216">Requirements</span></span>

|<span data-ttu-id="359f6-217">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-217">Requirement</span></span>| <span data-ttu-id="359f6-218">值</span><span class="sxs-lookup"><span data-stu-id="359f6-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-219">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-220">1.5</span><span class="sxs-lookup"><span data-stu-id="359f6-220">1.5</span></span> |
|[<span data-ttu-id="359f6-221">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-222">ReadItem</span></span> |
|[<span data-ttu-id="359f6-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-224">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-225">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-225">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="359f6-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="359f6-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="359f6-227">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="359f6-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-228">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="359f6-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="359f6-p107">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="359f6-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-231">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-231">Parameters</span></span>

|<span data-ttu-id="359f6-232">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-232">Name</span></span>| <span data-ttu-id="359f6-233">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-233">Type</span></span>| <span data-ttu-id="359f6-234">描述</span><span class="sxs-lookup"><span data-stu-id="359f6-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="359f6-235">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-235">String</span></span>|<span data-ttu-id="359f6-236">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="359f6-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="359f6-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="359f6-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="359f6-238">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="359f6-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-239">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-239">Requirements</span></span>

|<span data-ttu-id="359f6-240">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-240">Requirement</span></span>| <span data-ttu-id="359f6-241">值</span><span class="sxs-lookup"><span data-stu-id="359f6-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-242">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-243">1.3</span><span class="sxs-lookup"><span data-stu-id="359f6-243">1.3</span></span>|
|[<span data-ttu-id="359f6-244">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-245">受限</span><span class="sxs-lookup"><span data-stu-id="359f6-245">Restricted</span></span>|
|[<span data-ttu-id="359f6-246">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-247">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="359f6-248">返回：</span><span class="sxs-lookup"><span data-stu-id="359f6-248">Returns:</span></span>

<span data-ttu-id="359f6-249">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="359f6-250">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="359f6-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="359f6-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="359f6-252">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="359f6-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="359f6-p108">Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="359f6-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="359f6-p109">如果邮件应用程序在 Outlook 中运行，`convertToLocalClientTime` 方法将返回一个值设置为客户端计算机时区的字典对象。如果邮件应用程序在 Outlook Web App 中运行，`convertToLocalClientTime` 方法将返回值设置为 EAC 中指定的时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="359f6-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-258">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-258">Parameters</span></span>

|<span data-ttu-id="359f6-259">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-259">Name</span></span>| <span data-ttu-id="359f6-260">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-260">Type</span></span>| <span data-ttu-id="359f6-261">描述</span><span class="sxs-lookup"><span data-stu-id="359f6-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="359f6-262">日期</span><span class="sxs-lookup"><span data-stu-id="359f6-262">Date</span></span>|<span data-ttu-id="359f6-263">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="359f6-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-264">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-264">Requirements</span></span>

|<span data-ttu-id="359f6-265">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-265">Requirement</span></span>| <span data-ttu-id="359f6-266">值</span><span class="sxs-lookup"><span data-stu-id="359f6-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-268">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-268">1.0</span></span>|
|[<span data-ttu-id="359f6-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-270">ReadItem</span></span>|
|[<span data-ttu-id="359f6-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="359f6-273">返回：</span><span class="sxs-lookup"><span data-stu-id="359f6-273">Returns:</span></span>

<span data-ttu-id="359f6-274">类型：[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="359f6-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="359f6-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="359f6-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="359f6-276">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="359f6-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-277">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="359f6-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="359f6-p110">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="359f6-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-280">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-280">Parameters</span></span>

|<span data-ttu-id="359f6-281">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-281">Name</span></span>| <span data-ttu-id="359f6-282">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-282">Type</span></span>| <span data-ttu-id="359f6-283">描述</span><span class="sxs-lookup"><span data-stu-id="359f6-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="359f6-284">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-284">String</span></span>|<span data-ttu-id="359f6-285">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="359f6-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="359f6-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="359f6-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="359f6-287">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="359f6-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-288">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-288">Requirements</span></span>

|<span data-ttu-id="359f6-289">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-289">Requirement</span></span>| <span data-ttu-id="359f6-290">值</span><span class="sxs-lookup"><span data-stu-id="359f6-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-291">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-292">1.3</span><span class="sxs-lookup"><span data-stu-id="359f6-292">1.3</span></span>|
|[<span data-ttu-id="359f6-293">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-294">受限</span><span class="sxs-lookup"><span data-stu-id="359f6-294">Restricted</span></span>|
|[<span data-ttu-id="359f6-295">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-296">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="359f6-297">返回：</span><span class="sxs-lookup"><span data-stu-id="359f6-297">Returns:</span></span>

<span data-ttu-id="359f6-298">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="359f6-299">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="359f6-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="359f6-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="359f6-301">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="359f6-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="359f6-302">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="359f6-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-303">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-303">Parameters</span></span>

|<span data-ttu-id="359f6-304">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-304">Name</span></span>| <span data-ttu-id="359f6-305">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-305">Type</span></span>| <span data-ttu-id="359f6-306">说明</span><span class="sxs-lookup"><span data-stu-id="359f6-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="359f6-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="359f6-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="359f6-308">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="359f6-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-309">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-309">Requirements</span></span>

|<span data-ttu-id="359f6-310">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-310">Requirement</span></span>| <span data-ttu-id="359f6-311">值</span><span class="sxs-lookup"><span data-stu-id="359f6-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-312">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-313">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-313">1.0</span></span>|
|[<span data-ttu-id="359f6-314">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-315">ReadItem</span></span>|
|[<span data-ttu-id="359f6-316">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-317">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="359f6-318">返回：</span><span class="sxs-lookup"><span data-stu-id="359f6-318">Returns:</span></span>

<span data-ttu-id="359f6-319">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="359f6-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="359f6-320">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="359f6-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="359f6-321">日期</span><span class="sxs-lookup"><span data-stu-id="359f6-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="359f6-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="359f6-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="359f6-323">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="359f6-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-324">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="359f6-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="359f6-325">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="359f6-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="359f6-p111">在 Outlook for Mac 中，您可以使用此方法来显示不属于定期系列的单个约会，或显示定期系列的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。</span><span class="sxs-lookup"><span data-stu-id="359f6-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="359f6-328">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="359f6-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="359f6-329">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="359f6-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-330">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-330">Parameters</span></span>

|<span data-ttu-id="359f6-331">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-331">Name</span></span>| <span data-ttu-id="359f6-332">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-332">Type</span></span>| <span data-ttu-id="359f6-333">描述</span><span class="sxs-lookup"><span data-stu-id="359f6-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="359f6-334">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-334">String</span></span>|<span data-ttu-id="359f6-335">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="359f6-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-336">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-336">Requirements</span></span>

|<span data-ttu-id="359f6-337">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-337">Requirement</span></span>| <span data-ttu-id="359f6-338">值</span><span class="sxs-lookup"><span data-stu-id="359f6-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-339">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-340">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-340">1.0</span></span>|
|[<span data-ttu-id="359f6-341">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-342">ReadItem</span></span>|
|[<span data-ttu-id="359f6-343">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-344">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-345">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="359f6-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="359f6-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="359f6-347">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="359f6-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-348">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="359f6-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="359f6-349">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="359f6-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="359f6-350">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="359f6-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="359f6-351">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="359f6-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="359f6-p112">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="359f6-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-354">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-354">Parameters</span></span>

|<span data-ttu-id="359f6-355">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-355">Name</span></span>| <span data-ttu-id="359f6-356">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-356">Type</span></span>| <span data-ttu-id="359f6-357">描述</span><span class="sxs-lookup"><span data-stu-id="359f6-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="359f6-358">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-358">String</span></span>|<span data-ttu-id="359f6-359">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="359f6-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-360">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-360">Requirements</span></span>

|<span data-ttu-id="359f6-361">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-361">Requirement</span></span>| <span data-ttu-id="359f6-362">值</span><span class="sxs-lookup"><span data-stu-id="359f6-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-363">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-364">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-364">1.0</span></span>|
|[<span data-ttu-id="359f6-365">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-366">ReadItem</span></span>|
|[<span data-ttu-id="359f6-367">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-368">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-369">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="359f6-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="359f6-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="359f6-371">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="359f6-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-372">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="359f6-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="359f6-p113">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="359f6-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="359f6-p114">在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="359f6-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="359f6-p115">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="359f6-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="359f6-380">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="359f6-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-381">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-382">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="359f6-382">All parameters are optional.</span></span>

|<span data-ttu-id="359f6-383">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-383">Name</span></span>| <span data-ttu-id="359f6-384">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-384">Type</span></span>| <span data-ttu-id="359f6-385">描述</span><span class="sxs-lookup"><span data-stu-id="359f6-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="359f6-386">对象</span><span class="sxs-lookup"><span data-stu-id="359f6-386">Object</span></span> | <span data-ttu-id="359f6-387">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="359f6-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="359f6-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="359f6-p116">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="359f6-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="359f6-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="359f6-p117">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="359f6-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="359f6-394">Date</span><span class="sxs-lookup"><span data-stu-id="359f6-394">Date</span></span> | <span data-ttu-id="359f6-395">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="359f6-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="359f6-396">Date</span><span class="sxs-lookup"><span data-stu-id="359f6-396">Date</span></span> | <span data-ttu-id="359f6-397">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="359f6-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="359f6-398">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-398">String</span></span> | <span data-ttu-id="359f6-p118">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="359f6-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="359f6-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="359f6-p119">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="359f6-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="359f6-404">String</span><span class="sxs-lookup"><span data-stu-id="359f6-404">String</span></span> | <span data-ttu-id="359f6-p120">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="359f6-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="359f6-407">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-407">String</span></span> | <span data-ttu-id="359f6-p121">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="359f6-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="359f6-410">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-410">Requirements</span></span>

|<span data-ttu-id="359f6-411">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-411">Requirement</span></span>| <span data-ttu-id="359f6-412">值</span><span class="sxs-lookup"><span data-stu-id="359f6-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-413">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-414">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-414">1.0</span></span>|
|[<span data-ttu-id="359f6-415">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-416">ReadItem</span></span>|
|[<span data-ttu-id="359f6-417">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-418">阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-419">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="359f6-420">office.context.mailbox.displaynewmessageform (参数)</span><span class="sxs-lookup"><span data-stu-id="359f6-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="359f6-421">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="359f6-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="359f6-422">`displayNewMessageForm`方法将打开一个窗体, 使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="359f6-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="359f6-423">如果指定了参数, 则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="359f6-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="359f6-424">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="359f6-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-425">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-426">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="359f6-426">All parameters are optional.</span></span>

|<span data-ttu-id="359f6-427">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-427">Name</span></span>| <span data-ttu-id="359f6-428">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-428">Type</span></span>| <span data-ttu-id="359f6-429">描述</span><span class="sxs-lookup"><span data-stu-id="359f6-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="359f6-430">对象</span><span class="sxs-lookup"><span data-stu-id="359f6-430">Object</span></span> | <span data-ttu-id="359f6-431">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="359f6-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="359f6-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="359f6-433">包含电子邮件地址的字符串数组, 或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="359f6-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="359f6-434">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="359f6-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="359f6-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="359f6-436">包含电子邮件地址的字符串数组, 或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="359f6-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="359f6-437">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="359f6-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="359f6-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="359f6-439">包含电子邮件地址的字符串数组, 或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="359f6-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="359f6-440">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="359f6-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="359f6-441">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-441">String</span></span> | <span data-ttu-id="359f6-442">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="359f6-442">A string containing the subject of the message.</span></span> <span data-ttu-id="359f6-443">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="359f6-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="359f6-444">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-444">String</span></span> | <span data-ttu-id="359f6-445">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="359f6-445">The HTML body of the message.</span></span> <span data-ttu-id="359f6-446">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="359f6-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="359f6-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="359f6-448">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="359f6-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="359f6-449">String</span><span class="sxs-lookup"><span data-stu-id="359f6-449">String</span></span> | <span data-ttu-id="359f6-p128">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="359f6-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="359f6-452">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-452">String</span></span> | <span data-ttu-id="359f6-453">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="359f6-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="359f6-454">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-454">String</span></span> | <span data-ttu-id="359f6-p129">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="359f6-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="359f6-457">布尔</span><span class="sxs-lookup"><span data-stu-id="359f6-457">Boolean</span></span> | <span data-ttu-id="359f6-p130">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="359f6-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="359f6-460">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-460">String</span></span> | <span data-ttu-id="359f6-461">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="359f6-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="359f6-462">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="359f6-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="359f6-463">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="359f6-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="359f6-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-464">Requirements</span></span>

|<span data-ttu-id="359f6-465">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-465">Requirement</span></span>| <span data-ttu-id="359f6-466">值</span><span class="sxs-lookup"><span data-stu-id="359f6-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-468">1.6</span><span class="sxs-lookup"><span data-stu-id="359f6-468">1.6</span></span> |
|[<span data-ttu-id="359f6-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-470">ReadItem</span></span>|
|[<span data-ttu-id="359f6-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-472">阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-473">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-473">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="359f6-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="359f6-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="359f6-475">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="359f6-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="359f6-p132">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="359f6-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-478">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="359f6-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="359f6-479">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="359f6-479">**REST Tokens**</span></span>

<span data-ttu-id="359f6-p133">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="359f6-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="359f6-483">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="359f6-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="359f6-484">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="359f6-484">**EWS Tokens**</span></span>

<span data-ttu-id="359f6-p134">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="359f6-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="359f6-487">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="359f6-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-488">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-488">Parameters</span></span>

|<span data-ttu-id="359f6-489">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-489">Name</span></span>| <span data-ttu-id="359f6-490">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-490">Type</span></span>| <span data-ttu-id="359f6-491">属性</span><span class="sxs-lookup"><span data-stu-id="359f6-491">Attributes</span></span>| <span data-ttu-id="359f6-492">说明</span><span class="sxs-lookup"><span data-stu-id="359f6-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="359f6-493">Object</span><span class="sxs-lookup"><span data-stu-id="359f6-493">Object</span></span> | <span data-ttu-id="359f6-494">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-494">&lt;optional&gt;</span></span> | <span data-ttu-id="359f6-495">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="359f6-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="359f6-496">布尔值</span><span class="sxs-lookup"><span data-stu-id="359f6-496">Boolean</span></span> |  <span data-ttu-id="359f6-497">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-497">&lt;optional&gt;</span></span> | <span data-ttu-id="359f6-p135">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="359f6-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="359f6-500">Object</span><span class="sxs-lookup"><span data-stu-id="359f6-500">Object</span></span> |  <span data-ttu-id="359f6-501">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-501">&lt;optional&gt;</span></span> | <span data-ttu-id="359f6-502">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="359f6-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="359f6-503">函数</span><span class="sxs-lookup"><span data-stu-id="359f6-503">function</span></span>||<span data-ttu-id="359f6-p136">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="359f6-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-506">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-506">Requirements</span></span>

|<span data-ttu-id="359f6-507">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-507">Requirement</span></span>| <span data-ttu-id="359f6-508">值</span><span class="sxs-lookup"><span data-stu-id="359f6-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-509">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-510">1.5</span><span class="sxs-lookup"><span data-stu-id="359f6-510">1.5</span></span> |
|[<span data-ttu-id="359f6-511">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-512">ReadItem</span></span>|
|[<span data-ttu-id="359f6-513">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-514">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-515">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="359f6-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="359f6-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="359f6-517">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="359f6-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="359f6-p137">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="359f6-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="359f6-p138">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="359f6-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="359f6-523">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="359f6-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="359f6-p139">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="359f6-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-526">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-526">Parameters</span></span>

|<span data-ttu-id="359f6-527">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-527">Name</span></span>| <span data-ttu-id="359f6-528">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-528">Type</span></span>| <span data-ttu-id="359f6-529">属性</span><span class="sxs-lookup"><span data-stu-id="359f6-529">Attributes</span></span>| <span data-ttu-id="359f6-530">说明</span><span class="sxs-lookup"><span data-stu-id="359f6-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="359f6-531">函数</span><span class="sxs-lookup"><span data-stu-id="359f6-531">function</span></span>||<span data-ttu-id="359f6-p140">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="359f6-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="359f6-534">Object</span><span class="sxs-lookup"><span data-stu-id="359f6-534">Object</span></span>| <span data-ttu-id="359f6-535">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-535">&lt;optional&gt;</span></span>|<span data-ttu-id="359f6-536">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="359f6-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-537">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-537">Requirements</span></span>

|<span data-ttu-id="359f6-538">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-538">Requirement</span></span>| <span data-ttu-id="359f6-539">值</span><span class="sxs-lookup"><span data-stu-id="359f6-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-540">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-541">1.3</span><span class="sxs-lookup"><span data-stu-id="359f6-541">1.3</span></span>|
|[<span data-ttu-id="359f6-542">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-543">ReadItem</span></span>|
|[<span data-ttu-id="359f6-544">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-545">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-546">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-546">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="359f6-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="359f6-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="359f6-548">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="359f6-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="359f6-549">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="359f6-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-550">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-550">Parameters</span></span>

|<span data-ttu-id="359f6-551">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-551">Name</span></span>| <span data-ttu-id="359f6-552">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-552">Type</span></span>| <span data-ttu-id="359f6-553">属性</span><span class="sxs-lookup"><span data-stu-id="359f6-553">Attributes</span></span>| <span data-ttu-id="359f6-554">说明</span><span class="sxs-lookup"><span data-stu-id="359f6-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="359f6-555">function</span><span class="sxs-lookup"><span data-stu-id="359f6-555">function</span></span>||<span data-ttu-id="359f6-556">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="359f6-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="359f6-557">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="359f6-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="359f6-558">Object</span><span class="sxs-lookup"><span data-stu-id="359f6-558">Object</span></span>| <span data-ttu-id="359f6-559">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-559">&lt;optional&gt;</span></span>|<span data-ttu-id="359f6-560">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="359f6-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-561">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-561">Requirements</span></span>

|<span data-ttu-id="359f6-562">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-562">Requirement</span></span>| <span data-ttu-id="359f6-563">值</span><span class="sxs-lookup"><span data-stu-id="359f6-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-564">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-565">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-565">1.0</span></span>|
|[<span data-ttu-id="359f6-566">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-567">ReadItem</span></span>|
|[<span data-ttu-id="359f6-568">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-569">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-569">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-570">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-570">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="359f6-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="359f6-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="359f6-572">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="359f6-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-573">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="359f6-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="359f6-574">在 Outlook for iOS 或 Outlook for Android 中</span><span class="sxs-lookup"><span data-stu-id="359f6-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="359f6-575">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="359f6-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="359f6-576">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="359f6-576">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="359f6-577">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="359f6-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="359f6-578">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="359f6-578">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="359f6-579">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="359f6-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="359f6-580">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="359f6-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="359f6-p142">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="359f6-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="359f6-583">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="359f6-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="359f6-584">版本差异</span><span class="sxs-lookup"><span data-stu-id="359f6-584">Version differences</span></span>

<span data-ttu-id="359f6-585">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="359f6-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="359f6-p143">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="359f6-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-589">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-589">Parameters</span></span>

|<span data-ttu-id="359f6-590">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-590">Name</span></span>| <span data-ttu-id="359f6-591">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-591">Type</span></span>| <span data-ttu-id="359f6-592">属性</span><span class="sxs-lookup"><span data-stu-id="359f6-592">Attributes</span></span>| <span data-ttu-id="359f6-593">描述</span><span class="sxs-lookup"><span data-stu-id="359f6-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="359f6-594">字符串</span><span class="sxs-lookup"><span data-stu-id="359f6-594">String</span></span>||<span data-ttu-id="359f6-595">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="359f6-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="359f6-596">function</span><span class="sxs-lookup"><span data-stu-id="359f6-596">function</span></span>||<span data-ttu-id="359f6-597">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="359f6-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="359f6-598">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="359f6-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="359f6-599">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="359f6-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="359f6-600">对象</span><span class="sxs-lookup"><span data-stu-id="359f6-600">Object</span></span>| <span data-ttu-id="359f6-601">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-601">&lt;optional&gt;</span></span>|<span data-ttu-id="359f6-602">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="359f6-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-603">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-603">Requirements</span></span>

|<span data-ttu-id="359f6-604">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-604">Requirement</span></span>| <span data-ttu-id="359f6-605">值</span><span class="sxs-lookup"><span data-stu-id="359f6-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-606">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-607">1.0</span><span class="sxs-lookup"><span data-stu-id="359f6-607">1.0</span></span>|
|[<span data-ttu-id="359f6-608">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="359f6-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="359f6-610">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-611">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-611">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="359f6-612">示例</span><span class="sxs-lookup"><span data-stu-id="359f6-612">Example</span></span>

<span data-ttu-id="359f6-613">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="359f6-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="359f6-614">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="359f6-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="359f6-615">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="359f6-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="359f6-616">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="359f6-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="359f6-617">参数</span><span class="sxs-lookup"><span data-stu-id="359f6-617">Parameters</span></span>

| <span data-ttu-id="359f6-618">名称</span><span class="sxs-lookup"><span data-stu-id="359f6-618">Name</span></span> | <span data-ttu-id="359f6-619">类型</span><span class="sxs-lookup"><span data-stu-id="359f6-619">Type</span></span> | <span data-ttu-id="359f6-620">属性</span><span class="sxs-lookup"><span data-stu-id="359f6-620">Attributes</span></span> | <span data-ttu-id="359f6-621">说明</span><span class="sxs-lookup"><span data-stu-id="359f6-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="359f6-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="359f6-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="359f6-623">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="359f6-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="359f6-624">对象</span><span class="sxs-lookup"><span data-stu-id="359f6-624">Object</span></span> | <span data-ttu-id="359f6-625">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-625">&lt;optional&gt;</span></span> | <span data-ttu-id="359f6-626">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="359f6-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="359f6-627">Object</span><span class="sxs-lookup"><span data-stu-id="359f6-627">Object</span></span> | <span data-ttu-id="359f6-628">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-628">&lt;optional&gt;</span></span> | <span data-ttu-id="359f6-629">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="359f6-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="359f6-630">函数</span><span class="sxs-lookup"><span data-stu-id="359f6-630">function</span></span>| <span data-ttu-id="359f6-631">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="359f6-631">&lt;optional&gt;</span></span>|<span data-ttu-id="359f6-632">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="359f6-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="359f6-633">Requirements</span><span class="sxs-lookup"><span data-stu-id="359f6-633">Requirements</span></span>

|<span data-ttu-id="359f6-634">要求</span><span class="sxs-lookup"><span data-stu-id="359f6-634">Requirement</span></span>| <span data-ttu-id="359f6-635">值</span><span class="sxs-lookup"><span data-stu-id="359f6-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="359f6-636">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="359f6-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="359f6-637">1.5</span><span class="sxs-lookup"><span data-stu-id="359f6-637">1.5</span></span> |
|[<span data-ttu-id="359f6-638">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="359f6-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="359f6-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="359f6-639">ReadItem</span></span> |
|[<span data-ttu-id="359f6-640">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="359f6-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="359f6-641">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="359f6-641">Compose or Read</span></span>|
