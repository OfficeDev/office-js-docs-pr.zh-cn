---
title: Office.context.mailbox - 要求集 1.5
description: ''
ms.date: 11/27/2019
localization_priority: Priority
ms.openlocfilehash: eefeab2cf6fbe78451afae7e588640fe7f50dba4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629684"
---
# <a name="mailbox"></a><span data-ttu-id="42c43-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="42c43-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="42c43-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="42c43-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="42c43-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="42c43-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="42c43-105">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-105">Requirements</span></span>

|<span data-ttu-id="42c43-106">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-106">Requirement</span></span>| <span data-ttu-id="42c43-107">值</span><span class="sxs-lookup"><span data-stu-id="42c43-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-109">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-109">1.0</span></span>|
|[<span data-ttu-id="42c43-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-111">受限</span><span class="sxs-lookup"><span data-stu-id="42c43-111">Restricted</span></span>|
|[<span data-ttu-id="42c43-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="42c43-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="42c43-114">Members and methods</span></span>

| <span data-ttu-id="42c43-115">成员</span><span class="sxs-lookup"><span data-stu-id="42c43-115">Member</span></span> | <span data-ttu-id="42c43-116">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="42c43-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="42c43-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="42c43-118">成员</span><span class="sxs-lookup"><span data-stu-id="42c43-118">Member</span></span> |
| [<span data-ttu-id="42c43-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="42c43-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="42c43-120">成员</span><span class="sxs-lookup"><span data-stu-id="42c43-120">Member</span></span> |
| [<span data-ttu-id="42c43-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="42c43-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="42c43-122">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-122">Method</span></span> |
| [<span data-ttu-id="42c43-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="42c43-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="42c43-124">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-124">Method</span></span> |
| [<span data-ttu-id="42c43-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="42c43-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="42c43-126">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-126">Method</span></span> |
| [<span data-ttu-id="42c43-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="42c43-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="42c43-128">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-128">Method</span></span> |
| [<span data-ttu-id="42c43-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="42c43-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="42c43-130">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-130">Method</span></span> |
| [<span data-ttu-id="42c43-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="42c43-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="42c43-132">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-132">Method</span></span> |
| [<span data-ttu-id="42c43-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="42c43-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="42c43-134">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-134">Method</span></span> |
| [<span data-ttu-id="42c43-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="42c43-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="42c43-136">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-136">Method</span></span> |
| [<span data-ttu-id="42c43-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="42c43-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="42c43-138">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-138">Method</span></span> |
| [<span data-ttu-id="42c43-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="42c43-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="42c43-140">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-140">Method</span></span> |
| [<span data-ttu-id="42c43-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="42c43-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="42c43-142">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-142">Method</span></span> |
| [<span data-ttu-id="42c43-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="42c43-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="42c43-144">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-144">Method</span></span> |
| [<span data-ttu-id="42c43-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="42c43-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="42c43-146">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="42c43-147">命名空间</span><span class="sxs-lookup"><span data-stu-id="42c43-147">Namespaces</span></span>

<span data-ttu-id="42c43-148">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="42c43-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="42c43-149">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="42c43-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="42c43-150">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="42c43-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="42c43-151">Members</span><span class="sxs-lookup"><span data-stu-id="42c43-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="42c43-152">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="42c43-152">ewsUrl: String</span></span>

<span data-ttu-id="42c43-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="42c43-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-155">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="42c43-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="42c43-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="42c43-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="42c43-158">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="42c43-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="42c43-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="42c43-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="42c43-161">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-161">Type</span></span>

*   <span data-ttu-id="42c43-162">String</span><span class="sxs-lookup"><span data-stu-id="42c43-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="42c43-163">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-163">Requirements</span></span>

|<span data-ttu-id="42c43-164">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-164">Requirement</span></span>| <span data-ttu-id="42c43-165">值</span><span class="sxs-lookup"><span data-stu-id="42c43-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-166">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-167">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-167">1.0</span></span>|
|[<span data-ttu-id="42c43-168">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-169">ReadItem</span></span>|
|[<span data-ttu-id="42c43-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="42c43-172">restUrl：String</span><span class="sxs-lookup"><span data-stu-id="42c43-172">restUrl: String</span></span>

<span data-ttu-id="42c43-173">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="42c43-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="42c43-174">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="42c43-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-175">连接到配置了自定义 REST URL 的 Exchange 2016 或更高版本本地安装的 Outlook 客户端将返回 `restUrl` 的无效值。</span><span class="sxs-lookup"><span data-stu-id="42c43-175">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="42c43-176">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-176">Type</span></span>

*   <span data-ttu-id="42c43-177">String</span><span class="sxs-lookup"><span data-stu-id="42c43-177">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="42c43-178">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-178">Requirements</span></span>

|<span data-ttu-id="42c43-179">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-179">Requirement</span></span>| <span data-ttu-id="42c43-180">值</span><span class="sxs-lookup"><span data-stu-id="42c43-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-182">1.5</span><span class="sxs-lookup"><span data-stu-id="42c43-182">1.5</span></span> |
|[<span data-ttu-id="42c43-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-184">ReadItem</span></span>|
|[<span data-ttu-id="42c43-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-186">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="42c43-187">方法</span><span class="sxs-lookup"><span data-stu-id="42c43-187">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="42c43-188">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="42c43-188">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="42c43-189">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="42c43-189">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="42c43-190">目前，唯一受支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该类型。</span><span class="sxs-lookup"><span data-stu-id="42c43-190">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="42c43-191">此事件由实现可固定任务窗格的加载项使用，并允许加载项根据当前选定的项目刷新任务窗格 UI。</span><span class="sxs-lookup"><span data-stu-id="42c43-191">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-192">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-192">Parameters</span></span>

| <span data-ttu-id="42c43-193">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-193">Name</span></span> | <span data-ttu-id="42c43-194">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-194">Type</span></span> | <span data-ttu-id="42c43-195">属性</span><span class="sxs-lookup"><span data-stu-id="42c43-195">Attributes</span></span> | <span data-ttu-id="42c43-196">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="42c43-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="42c43-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="42c43-198">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="42c43-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="42c43-199">函数</span><span class="sxs-lookup"><span data-stu-id="42c43-199">Function</span></span> || <span data-ttu-id="42c43-p105">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="42c43-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="42c43-203">Object</span><span class="sxs-lookup"><span data-stu-id="42c43-203">Object</span></span> | <span data-ttu-id="42c43-204">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-204">&lt;optional&gt;</span></span> | <span data-ttu-id="42c43-205">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="42c43-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="42c43-206">对象</span><span class="sxs-lookup"><span data-stu-id="42c43-206">Object</span></span> | <span data-ttu-id="42c43-207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-207">&lt;optional&gt;</span></span> | <span data-ttu-id="42c43-208">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="42c43-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="42c43-209">函数</span><span class="sxs-lookup"><span data-stu-id="42c43-209">function</span></span>| <span data-ttu-id="42c43-210">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-210">&lt;optional&gt;</span></span>|<span data-ttu-id="42c43-211">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="42c43-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="42c43-212">Requirements</span></span>

|<span data-ttu-id="42c43-213">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-213">Requirement</span></span>| <span data-ttu-id="42c43-214">值</span><span class="sxs-lookup"><span data-stu-id="42c43-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-215">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-216">1.5</span><span class="sxs-lookup"><span data-stu-id="42c43-216">1.5</span></span> |
|[<span data-ttu-id="42c43-217">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-218">ReadItem</span></span> |
|[<span data-ttu-id="42c43-219">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-220">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="42c43-221">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-221">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="42c43-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="42c43-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="42c43-223">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="42c43-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-224">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="42c43-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="42c43-p106">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="42c43-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-227">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-227">Parameters</span></span>

|<span data-ttu-id="42c43-228">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-228">Name</span></span>| <span data-ttu-id="42c43-229">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-229">Type</span></span>| <span data-ttu-id="42c43-230">描述</span><span class="sxs-lookup"><span data-stu-id="42c43-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="42c43-231">字符串</span><span class="sxs-lookup"><span data-stu-id="42c43-231">String</span></span>|<span data-ttu-id="42c43-232">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="42c43-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="42c43-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="42c43-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="42c43-234">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="42c43-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-235">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-235">Requirements</span></span>

|<span data-ttu-id="42c43-236">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-236">Requirement</span></span>| <span data-ttu-id="42c43-237">值</span><span class="sxs-lookup"><span data-stu-id="42c43-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-238">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-239">1.3</span><span class="sxs-lookup"><span data-stu-id="42c43-239">1.3</span></span>|
|[<span data-ttu-id="42c43-240">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-241">受限</span><span class="sxs-lookup"><span data-stu-id="42c43-241">Restricted</span></span>|
|[<span data-ttu-id="42c43-242">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-243">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="42c43-244">返回：</span><span class="sxs-lookup"><span data-stu-id="42c43-244">Returns:</span></span>

<span data-ttu-id="42c43-245">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="42c43-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="42c43-246">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="42c43-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="42c43-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="42c43-248">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="42c43-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="42c43-p107">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="42c43-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="42c43-p108">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="42c43-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-254">Parameters</span><span class="sxs-lookup"><span data-stu-id="42c43-254">Parameters</span></span>

|<span data-ttu-id="42c43-255">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-255">Name</span></span>| <span data-ttu-id="42c43-256">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-256">Type</span></span>| <span data-ttu-id="42c43-257">描述</span><span class="sxs-lookup"><span data-stu-id="42c43-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="42c43-258">日期</span><span class="sxs-lookup"><span data-stu-id="42c43-258">Date</span></span>|<span data-ttu-id="42c43-259">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="42c43-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-260">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-260">Requirements</span></span>

|<span data-ttu-id="42c43-261">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-261">Requirement</span></span>| <span data-ttu-id="42c43-262">值</span><span class="sxs-lookup"><span data-stu-id="42c43-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-264">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-264">1.0</span></span>|
|[<span data-ttu-id="42c43-265">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-266">ReadItem</span></span>|
|[<span data-ttu-id="42c43-267">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-268">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="42c43-269">返回：</span><span class="sxs-lookup"><span data-stu-id="42c43-269">Returns:</span></span>

<span data-ttu-id="42c43-270">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="42c43-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="42c43-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="42c43-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="42c43-272">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="42c43-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-273">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="42c43-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="42c43-p109">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="42c43-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-276">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-276">Parameters</span></span>

|<span data-ttu-id="42c43-277">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-277">Name</span></span>| <span data-ttu-id="42c43-278">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-278">Type</span></span>| <span data-ttu-id="42c43-279">描述</span><span class="sxs-lookup"><span data-stu-id="42c43-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="42c43-280">字符串</span><span class="sxs-lookup"><span data-stu-id="42c43-280">String</span></span>|<span data-ttu-id="42c43-281">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="42c43-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="42c43-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="42c43-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="42c43-283">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="42c43-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-284">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-284">Requirements</span></span>

|<span data-ttu-id="42c43-285">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-285">Requirement</span></span>| <span data-ttu-id="42c43-286">值</span><span class="sxs-lookup"><span data-stu-id="42c43-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-287">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-288">1.3</span><span class="sxs-lookup"><span data-stu-id="42c43-288">1.3</span></span>|
|[<span data-ttu-id="42c43-289">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-290">受限</span><span class="sxs-lookup"><span data-stu-id="42c43-290">Restricted</span></span>|
|[<span data-ttu-id="42c43-291">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-292">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="42c43-293">返回：</span><span class="sxs-lookup"><span data-stu-id="42c43-293">Returns:</span></span>

<span data-ttu-id="42c43-294">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="42c43-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="42c43-295">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="42c43-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="42c43-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="42c43-297">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="42c43-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="42c43-298">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="42c43-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-299">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-299">Parameters</span></span>

|<span data-ttu-id="42c43-300">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-300">Name</span></span>| <span data-ttu-id="42c43-301">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-301">Type</span></span>| <span data-ttu-id="42c43-302">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="42c43-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="42c43-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="42c43-304">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="42c43-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-305">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-305">Requirements</span></span>

|<span data-ttu-id="42c43-306">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-306">Requirement</span></span>| <span data-ttu-id="42c43-307">值</span><span class="sxs-lookup"><span data-stu-id="42c43-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-309">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-309">1.0</span></span>|
|[<span data-ttu-id="42c43-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-311">ReadItem</span></span>|
|[<span data-ttu-id="42c43-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-313">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="42c43-314">返回：</span><span class="sxs-lookup"><span data-stu-id="42c43-314">Returns:</span></span>

<span data-ttu-id="42c43-315">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="42c43-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="42c43-316">键入：日期</span><span class="sxs-lookup"><span data-stu-id="42c43-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="42c43-317">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-317">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="42c43-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="42c43-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="42c43-319">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="42c43-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-320">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="42c43-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="42c43-321">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="42c43-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="42c43-p110">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="42c43-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="42c43-324">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="42c43-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="42c43-325">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="42c43-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-326">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-326">Parameters</span></span>

|<span data-ttu-id="42c43-327">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-327">Name</span></span>| <span data-ttu-id="42c43-328">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-328">Type</span></span>| <span data-ttu-id="42c43-329">描述</span><span class="sxs-lookup"><span data-stu-id="42c43-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="42c43-330">字符串</span><span class="sxs-lookup"><span data-stu-id="42c43-330">String</span></span>|<span data-ttu-id="42c43-331">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="42c43-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-332">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-332">Requirements</span></span>

|<span data-ttu-id="42c43-333">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-333">Requirement</span></span>| <span data-ttu-id="42c43-334">值</span><span class="sxs-lookup"><span data-stu-id="42c43-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-335">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-336">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-336">1.0</span></span>|
|[<span data-ttu-id="42c43-337">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-338">ReadItem</span></span>|
|[<span data-ttu-id="42c43-339">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-340">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="42c43-341">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="42c43-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="42c43-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="42c43-343">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="42c43-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-344">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="42c43-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="42c43-345">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="42c43-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="42c43-346">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="42c43-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="42c43-347">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="42c43-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="42c43-p111">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="42c43-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-350">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-350">Parameters</span></span>

|<span data-ttu-id="42c43-351">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-351">Name</span></span>| <span data-ttu-id="42c43-352">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-352">Type</span></span>| <span data-ttu-id="42c43-353">描述</span><span class="sxs-lookup"><span data-stu-id="42c43-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="42c43-354">字符串</span><span class="sxs-lookup"><span data-stu-id="42c43-354">String</span></span>|<span data-ttu-id="42c43-355">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="42c43-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-356">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-356">Requirements</span></span>

|<span data-ttu-id="42c43-357">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-357">Requirement</span></span>| <span data-ttu-id="42c43-358">值</span><span class="sxs-lookup"><span data-stu-id="42c43-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-359">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-360">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-360">1.0</span></span>|
|[<span data-ttu-id="42c43-361">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-362">ReadItem</span></span>|
|[<span data-ttu-id="42c43-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-364">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="42c43-365">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="42c43-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="42c43-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="42c43-367">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="42c43-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-368">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="42c43-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="42c43-p112">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="42c43-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="42c43-p113">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="42c43-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="42c43-p114">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="42c43-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="42c43-376">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="42c43-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-377">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-377">Parameters</span></span>

|<span data-ttu-id="42c43-378">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-378">Name</span></span>| <span data-ttu-id="42c43-379">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-379">Type</span></span>| <span data-ttu-id="42c43-380">描述</span><span class="sxs-lookup"><span data-stu-id="42c43-380">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="42c43-381">对象</span><span class="sxs-lookup"><span data-stu-id="42c43-381">Object</span></span> | <span data-ttu-id="42c43-382">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="42c43-382">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="42c43-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="42c43-p115">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="42c43-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="42c43-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="42c43-p116">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="42c43-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="42c43-389">日期</span><span class="sxs-lookup"><span data-stu-id="42c43-389">Date</span></span> | <span data-ttu-id="42c43-390">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="42c43-390">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="42c43-391">Date</span><span class="sxs-lookup"><span data-stu-id="42c43-391">Date</span></span> | <span data-ttu-id="42c43-392">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="42c43-392">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="42c43-393">String</span><span class="sxs-lookup"><span data-stu-id="42c43-393">String</span></span> | <span data-ttu-id="42c43-p117">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="42c43-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="42c43-396">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-396">Array.&lt;String&gt;</span></span> | <span data-ttu-id="42c43-p118">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="42c43-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="42c43-399">String</span><span class="sxs-lookup"><span data-stu-id="42c43-399">String</span></span> | <span data-ttu-id="42c43-p119">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="42c43-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="42c43-402">字符串</span><span class="sxs-lookup"><span data-stu-id="42c43-402">String</span></span> | <span data-ttu-id="42c43-p120">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="42c43-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="42c43-405">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-405">Requirements</span></span>

|<span data-ttu-id="42c43-406">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-406">Requirement</span></span>| <span data-ttu-id="42c43-407">值</span><span class="sxs-lookup"><span data-stu-id="42c43-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-408">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-409">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-409">1.0</span></span>|
|[<span data-ttu-id="42c43-410">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-411">ReadItem</span></span>|
|[<span data-ttu-id="42c43-412">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-413">阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="42c43-414">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-414">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="42c43-415">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="42c43-415">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="42c43-416">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="42c43-416">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="42c43-p121">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="42c43-p121">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-419">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="42c43-419">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="42c43-420">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="42c43-420">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="42c43-421">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="42c43-421">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="42c43-422">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="42c43-422">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="42c43-423">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="42c43-423">**REST Tokens**</span></span>

<span data-ttu-id="42c43-p123">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="42c43-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="42c43-427">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="42c43-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="42c43-428">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="42c43-428">**EWS Tokens**</span></span>

<span data-ttu-id="42c43-p124">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="42c43-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="42c43-431">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="42c43-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="42c43-432">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="42c43-432">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="42c43-433">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="42c43-433">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="42c43-434">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="42c43-434">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-435">Parameters</span><span class="sxs-lookup"><span data-stu-id="42c43-435">Parameters</span></span>

|<span data-ttu-id="42c43-436">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-436">Name</span></span>| <span data-ttu-id="42c43-437">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-437">Type</span></span>| <span data-ttu-id="42c43-438">属性</span><span class="sxs-lookup"><span data-stu-id="42c43-438">Attributes</span></span>| <span data-ttu-id="42c43-439">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-439">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="42c43-440">Object</span><span class="sxs-lookup"><span data-stu-id="42c43-440">Object</span></span> | <span data-ttu-id="42c43-441">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-441">&lt;optional&gt;</span></span> | <span data-ttu-id="42c43-442">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="42c43-442">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="42c43-443">布尔值</span><span class="sxs-lookup"><span data-stu-id="42c43-443">Boolean</span></span> |  <span data-ttu-id="42c43-444">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-444">&lt;optional&gt;</span></span> | <span data-ttu-id="42c43-p126">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="42c43-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="42c43-447">Object</span><span class="sxs-lookup"><span data-stu-id="42c43-447">Object</span></span> |  <span data-ttu-id="42c43-448">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-448">&lt;optional&gt;</span></span> | <span data-ttu-id="42c43-449">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="42c43-449">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="42c43-450">函数</span><span class="sxs-lookup"><span data-stu-id="42c43-450">function</span></span>||<span data-ttu-id="42c43-451">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="42c43-451">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="42c43-452">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="42c43-452">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="42c43-453">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="42c43-453">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="42c43-454">错误</span><span class="sxs-lookup"><span data-stu-id="42c43-454">Errors</span></span>

|<span data-ttu-id="42c43-455">错误代码</span><span class="sxs-lookup"><span data-stu-id="42c43-455">Error code</span></span>|<span data-ttu-id="42c43-456">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-456">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="42c43-457">请求失败。</span><span class="sxs-lookup"><span data-stu-id="42c43-457">The request has failed.</span></span> <span data-ttu-id="42c43-458">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="42c43-458">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="42c43-459">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="42c43-459">The Exchange server returned an error.</span></span> <span data-ttu-id="42c43-460">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="42c43-460">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="42c43-461">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="42c43-461">The user is no longer connected to the network.</span></span> <span data-ttu-id="42c43-462">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="42c43-462">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-463">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-463">Requirements</span></span>

|<span data-ttu-id="42c43-464">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-464">Requirement</span></span>| <span data-ttu-id="42c43-465">值</span><span class="sxs-lookup"><span data-stu-id="42c43-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-466">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-467">1.5</span><span class="sxs-lookup"><span data-stu-id="42c43-467">1.5</span></span> |
|[<span data-ttu-id="42c43-468">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-469">ReadItem</span></span>|
|[<span data-ttu-id="42c43-470">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-471">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-471">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="42c43-472">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-472">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="42c43-473">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="42c43-473">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="42c43-474">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="42c43-474">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="42c43-p130">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="42c43-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="42c43-477">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="42c43-477">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="42c43-478">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="42c43-478">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="42c43-479">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="42c43-479">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="42c43-480">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="42c43-480">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="42c43-481">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="42c43-481">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="42c43-482">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="42c43-482">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-483">Parameters</span><span class="sxs-lookup"><span data-stu-id="42c43-483">Parameters</span></span>

|<span data-ttu-id="42c43-484">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-484">Name</span></span>| <span data-ttu-id="42c43-485">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-485">Type</span></span>| <span data-ttu-id="42c43-486">属性</span><span class="sxs-lookup"><span data-stu-id="42c43-486">Attributes</span></span>| <span data-ttu-id="42c43-487">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-487">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="42c43-488">函数</span><span class="sxs-lookup"><span data-stu-id="42c43-488">function</span></span>||<span data-ttu-id="42c43-489">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="42c43-489">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="42c43-490">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="42c43-490">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="42c43-491">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="42c43-491">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="42c43-492">对象</span><span class="sxs-lookup"><span data-stu-id="42c43-492">Object</span></span>| <span data-ttu-id="42c43-493">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-493">&lt;optional&gt;</span></span>|<span data-ttu-id="42c43-494">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="42c43-494">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="42c43-495">错误</span><span class="sxs-lookup"><span data-stu-id="42c43-495">Errors</span></span>

|<span data-ttu-id="42c43-496">错误代码</span><span class="sxs-lookup"><span data-stu-id="42c43-496">Error code</span></span>|<span data-ttu-id="42c43-497">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-497">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="42c43-498">请求失败。</span><span class="sxs-lookup"><span data-stu-id="42c43-498">The request has failed.</span></span> <span data-ttu-id="42c43-499">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="42c43-499">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="42c43-500">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="42c43-500">The Exchange server returned an error.</span></span> <span data-ttu-id="42c43-501">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="42c43-501">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="42c43-502">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="42c43-502">The user is no longer connected to the network.</span></span> <span data-ttu-id="42c43-503">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="42c43-503">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-504">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-504">Requirements</span></span>

|<span data-ttu-id="42c43-505">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-505">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="42c43-506">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-507">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-507">1.0</span></span> | <span data-ttu-id="42c43-508">1.3</span><span class="sxs-lookup"><span data-stu-id="42c43-508">1.3</span></span> |
|[<span data-ttu-id="42c43-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-510">ReadItem</span></span> | <span data-ttu-id="42c43-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-511">ReadItem</span></span> |
|[<span data-ttu-id="42c43-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-513">阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-513">Read</span></span> | <span data-ttu-id="42c43-514">撰写</span><span class="sxs-lookup"><span data-stu-id="42c43-514">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="42c43-515">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-515">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="42c43-516">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="42c43-516">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="42c43-517">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="42c43-517">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="42c43-518">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="42c43-518">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-519">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-519">Parameters</span></span>

|<span data-ttu-id="42c43-520">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-520">Name</span></span>| <span data-ttu-id="42c43-521">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-521">Type</span></span>| <span data-ttu-id="42c43-522">属性</span><span class="sxs-lookup"><span data-stu-id="42c43-522">Attributes</span></span>| <span data-ttu-id="42c43-523">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-523">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="42c43-524">函数</span><span class="sxs-lookup"><span data-stu-id="42c43-524">function</span></span>||<span data-ttu-id="42c43-525">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="42c43-525">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="42c43-526">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="42c43-526">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="42c43-527">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="42c43-527">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="42c43-528">对象</span><span class="sxs-lookup"><span data-stu-id="42c43-528">Object</span></span>| <span data-ttu-id="42c43-529">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-529">&lt;optional&gt;</span></span>|<span data-ttu-id="42c43-530">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="42c43-530">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="42c43-531">错误</span><span class="sxs-lookup"><span data-stu-id="42c43-531">Errors</span></span>

|<span data-ttu-id="42c43-532">错误代码</span><span class="sxs-lookup"><span data-stu-id="42c43-532">Error code</span></span>|<span data-ttu-id="42c43-533">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-533">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="42c43-534">请求失败。</span><span class="sxs-lookup"><span data-stu-id="42c43-534">The request has failed.</span></span> <span data-ttu-id="42c43-535">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="42c43-535">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="42c43-536">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="42c43-536">The Exchange server returned an error.</span></span> <span data-ttu-id="42c43-537">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="42c43-537">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="42c43-538">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="42c43-538">The user is no longer connected to the network.</span></span> <span data-ttu-id="42c43-539">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="42c43-539">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-540">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-540">Requirements</span></span>

|<span data-ttu-id="42c43-541">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-541">Requirement</span></span>| <span data-ttu-id="42c43-542">值</span><span class="sxs-lookup"><span data-stu-id="42c43-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-543">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-544">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-544">1.0</span></span>|
|[<span data-ttu-id="42c43-545">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-545">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-546">ReadItem</span></span>|
|[<span data-ttu-id="42c43-547">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-547">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-548">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-548">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="42c43-549">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-549">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="42c43-550">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="42c43-550">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="42c43-551">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="42c43-551">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-552">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="42c43-552">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="42c43-553">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="42c43-553">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="42c43-554">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="42c43-554">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="42c43-555">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="42c43-555">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="42c43-556">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="42c43-556">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="42c43-557">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="42c43-557">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="42c43-558">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="42c43-558">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="42c43-559">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="42c43-559">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="42c43-p140">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="42c43-p140">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="42c43-562">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="42c43-562">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="42c43-563">版本差异</span><span class="sxs-lookup"><span data-stu-id="42c43-563">Version differences</span></span>

<span data-ttu-id="42c43-564">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="42c43-564">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="42c43-p141">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="42c43-p141">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-568">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-568">Parameters</span></span>

|<span data-ttu-id="42c43-569">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-569">Name</span></span>| <span data-ttu-id="42c43-570">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-570">Type</span></span>| <span data-ttu-id="42c43-571">属性</span><span class="sxs-lookup"><span data-stu-id="42c43-571">Attributes</span></span>| <span data-ttu-id="42c43-572">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-572">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="42c43-573">字符串</span><span class="sxs-lookup"><span data-stu-id="42c43-573">String</span></span>||<span data-ttu-id="42c43-574">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="42c43-574">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="42c43-575">函数</span><span class="sxs-lookup"><span data-stu-id="42c43-575">function</span></span>||<span data-ttu-id="42c43-576">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="42c43-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="42c43-577">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="42c43-577">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="42c43-578">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="42c43-578">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="42c43-579">对象</span><span class="sxs-lookup"><span data-stu-id="42c43-579">Object</span></span>| <span data-ttu-id="42c43-580">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-580">&lt;optional&gt;</span></span>|<span data-ttu-id="42c43-581">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="42c43-581">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-582">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-582">Requirements</span></span>

|<span data-ttu-id="42c43-583">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-583">Requirement</span></span>| <span data-ttu-id="42c43-584">值</span><span class="sxs-lookup"><span data-stu-id="42c43-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-585">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-586">1.0</span><span class="sxs-lookup"><span data-stu-id="42c43-586">1.0</span></span>|
|[<span data-ttu-id="42c43-587">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-587">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-588">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="42c43-588">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="42c43-589">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-589">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-590">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-590">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="42c43-591">示例</span><span class="sxs-lookup"><span data-stu-id="42c43-591">Example</span></span>

<span data-ttu-id="42c43-592">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="42c43-592">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="42c43-593">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="42c43-593">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="42c43-594">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="42c43-594">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="42c43-595">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="42c43-595">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="42c43-596">参数</span><span class="sxs-lookup"><span data-stu-id="42c43-596">Parameters</span></span>

| <span data-ttu-id="42c43-597">名称</span><span class="sxs-lookup"><span data-stu-id="42c43-597">Name</span></span> | <span data-ttu-id="42c43-598">类型</span><span class="sxs-lookup"><span data-stu-id="42c43-598">Type</span></span> | <span data-ttu-id="42c43-599">属性</span><span class="sxs-lookup"><span data-stu-id="42c43-599">Attributes</span></span> | <span data-ttu-id="42c43-600">说明</span><span class="sxs-lookup"><span data-stu-id="42c43-600">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="42c43-601">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="42c43-601">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="42c43-602">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="42c43-602">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="42c43-603">对象</span><span class="sxs-lookup"><span data-stu-id="42c43-603">Object</span></span> | <span data-ttu-id="42c43-604">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-604">&lt;optional&gt;</span></span> | <span data-ttu-id="42c43-605">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="42c43-605">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="42c43-606">对象</span><span class="sxs-lookup"><span data-stu-id="42c43-606">Object</span></span> | <span data-ttu-id="42c43-607">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-607">&lt;optional&gt;</span></span> | <span data-ttu-id="42c43-608">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="42c43-608">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="42c43-609">函数</span><span class="sxs-lookup"><span data-stu-id="42c43-609">function</span></span>| <span data-ttu-id="42c43-610">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="42c43-610">&lt;optional&gt;</span></span>|<span data-ttu-id="42c43-611">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="42c43-611">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="42c43-612">Requirements</span><span class="sxs-lookup"><span data-stu-id="42c43-612">Requirements</span></span>

|<span data-ttu-id="42c43-613">要求</span><span class="sxs-lookup"><span data-stu-id="42c43-613">Requirement</span></span>| <span data-ttu-id="42c43-614">值</span><span class="sxs-lookup"><span data-stu-id="42c43-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="42c43-615">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="42c43-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="42c43-616">1.5</span><span class="sxs-lookup"><span data-stu-id="42c43-616">1.5</span></span> |
|[<span data-ttu-id="42c43-617">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="42c43-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="42c43-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="42c43-618">ReadItem</span></span> |
|[<span data-ttu-id="42c43-619">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="42c43-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="42c43-620">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="42c43-620">Compose or Read</span></span>|
