---
title: Office.context.mailbox - 要求集 1.5
description: ''
ms.date: 10/21/2019
localization_priority: Priority
ms.openlocfilehash: bb63d8186d41d072aa62b180b16958d61ce9a66c
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627011"
---
# <a name="mailbox"></a><span data-ttu-id="d4854-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="d4854-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="d4854-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="d4854-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="d4854-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d4854-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4854-105">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-105">Requirements</span></span>

|<span data-ttu-id="d4854-106">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-106">Requirement</span></span>| <span data-ttu-id="d4854-107">值</span><span class="sxs-lookup"><span data-stu-id="d4854-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-109">1.0</span></span>|
|[<span data-ttu-id="d4854-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-111">受限</span><span class="sxs-lookup"><span data-stu-id="d4854-111">Restricted</span></span>|
|[<span data-ttu-id="d4854-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d4854-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="d4854-114">Members and methods</span></span>

| <span data-ttu-id="d4854-115">成员</span><span class="sxs-lookup"><span data-stu-id="d4854-115">Member</span></span> | <span data-ttu-id="d4854-116">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d4854-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d4854-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d4854-118">成员</span><span class="sxs-lookup"><span data-stu-id="d4854-118">Member</span></span> |
| [<span data-ttu-id="d4854-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="d4854-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d4854-120">成员</span><span class="sxs-lookup"><span data-stu-id="d4854-120">Member</span></span> |
| [<span data-ttu-id="d4854-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d4854-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d4854-122">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-122">Method</span></span> |
| [<span data-ttu-id="d4854-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d4854-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d4854-124">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-124">Method</span></span> |
| [<span data-ttu-id="d4854-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d4854-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="d4854-126">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-126">Method</span></span> |
| [<span data-ttu-id="d4854-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d4854-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d4854-128">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-128">Method</span></span> |
| [<span data-ttu-id="d4854-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d4854-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d4854-130">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-130">Method</span></span> |
| [<span data-ttu-id="d4854-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d4854-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d4854-132">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-132">Method</span></span> |
| [<span data-ttu-id="d4854-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d4854-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d4854-134">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-134">Method</span></span> |
| [<span data-ttu-id="d4854-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d4854-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d4854-136">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-136">Method</span></span> |
| [<span data-ttu-id="d4854-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d4854-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d4854-138">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-138">Method</span></span> |
| [<span data-ttu-id="d4854-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d4854-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d4854-140">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-140">Method</span></span> |
| [<span data-ttu-id="d4854-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d4854-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d4854-142">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-142">Method</span></span> |
| [<span data-ttu-id="d4854-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d4854-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d4854-144">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-144">Method</span></span> |
| [<span data-ttu-id="d4854-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d4854-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d4854-146">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d4854-147">命名空间</span><span class="sxs-lookup"><span data-stu-id="d4854-147">Namespaces</span></span>

<span data-ttu-id="d4854-148">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="d4854-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d4854-149">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="d4854-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d4854-150">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="d4854-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d4854-151">Members</span><span class="sxs-lookup"><span data-stu-id="d4854-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d4854-152">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="d4854-152">ewsUrl: String</span></span>

<span data-ttu-id="d4854-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d4854-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-155">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="d4854-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d4854-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="d4854-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d4854-158">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="d4854-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d4854-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="d4854-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d4854-161">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-161">Type</span></span>

*   <span data-ttu-id="d4854-162">String</span><span class="sxs-lookup"><span data-stu-id="d4854-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4854-163">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-163">Requirements</span></span>

|<span data-ttu-id="d4854-164">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-164">Requirement</span></span>| <span data-ttu-id="d4854-165">值</span><span class="sxs-lookup"><span data-stu-id="d4854-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-166">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-167">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-167">1.0</span></span>|
|[<span data-ttu-id="d4854-168">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-169">ReadItem</span></span>|
|[<span data-ttu-id="d4854-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="d4854-172">restUrl：String</span><span class="sxs-lookup"><span data-stu-id="d4854-172">restUrl: String</span></span>

<span data-ttu-id="d4854-173">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="d4854-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d4854-174">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="d4854-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="d4854-175">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="d4854-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="d4854-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="d4854-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-178">连接到配置了自定义 REST URL 的 Exchange 2016 或更高版本本地安装的 Outlook 客户端将返回 `restUrl` 的无效值。</span><span class="sxs-lookup"><span data-stu-id="d4854-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="d4854-179">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-179">Type</span></span>

*   <span data-ttu-id="d4854-180">String</span><span class="sxs-lookup"><span data-stu-id="d4854-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4854-181">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-181">Requirements</span></span>

|<span data-ttu-id="d4854-182">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-182">Requirement</span></span>| <span data-ttu-id="d4854-183">值</span><span class="sxs-lookup"><span data-stu-id="d4854-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-185">1.5</span><span class="sxs-lookup"><span data-stu-id="d4854-185">1.5</span></span> |
|[<span data-ttu-id="d4854-186">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-187">ReadItem</span></span>|
|[<span data-ttu-id="d4854-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d4854-190">方法</span><span class="sxs-lookup"><span data-stu-id="d4854-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d4854-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d4854-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d4854-192">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d4854-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d4854-193">目前，唯一受支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该类型。</span><span class="sxs-lookup"><span data-stu-id="d4854-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="d4854-194">此事件由实现可固定任务窗格的加载项使用，并允许加载项根据当前选定的项目刷新任务窗格 UI。</span><span class="sxs-lookup"><span data-stu-id="d4854-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-195">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-195">Parameters</span></span>

| <span data-ttu-id="d4854-196">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-196">Name</span></span> | <span data-ttu-id="d4854-197">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-197">Type</span></span> | <span data-ttu-id="d4854-198">属性</span><span class="sxs-lookup"><span data-stu-id="d4854-198">Attributes</span></span> | <span data-ttu-id="d4854-199">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d4854-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d4854-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d4854-201">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="d4854-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d4854-202">函数</span><span class="sxs-lookup"><span data-stu-id="d4854-202">Function</span></span> || <span data-ttu-id="d4854-p106">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="d4854-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d4854-206">Object</span><span class="sxs-lookup"><span data-stu-id="d4854-206">Object</span></span> | <span data-ttu-id="d4854-207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-207">&lt;optional&gt;</span></span> | <span data-ttu-id="d4854-208">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d4854-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d4854-209">对象</span><span class="sxs-lookup"><span data-stu-id="d4854-209">Object</span></span> | <span data-ttu-id="d4854-210">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-210">&lt;optional&gt;</span></span> | <span data-ttu-id="d4854-211">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d4854-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d4854-212">函数</span><span class="sxs-lookup"><span data-stu-id="d4854-212">function</span></span>| <span data-ttu-id="d4854-213">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-213">&lt;optional&gt;</span></span>|<span data-ttu-id="d4854-214">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d4854-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-215">Requirements</span><span class="sxs-lookup"><span data-stu-id="d4854-215">Requirements</span></span>

|<span data-ttu-id="d4854-216">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-216">Requirement</span></span>| <span data-ttu-id="d4854-217">值</span><span class="sxs-lookup"><span data-stu-id="d4854-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-219">1.5</span><span class="sxs-lookup"><span data-stu-id="d4854-219">1.5</span></span> |
|[<span data-ttu-id="d4854-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-221">ReadItem</span></span> |
|[<span data-ttu-id="d4854-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-223">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4854-224">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d4854-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d4854-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d4854-226">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="d4854-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-227">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d4854-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d4854-p107">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="d4854-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-230">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-230">Parameters</span></span>

|<span data-ttu-id="d4854-231">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-231">Name</span></span>| <span data-ttu-id="d4854-232">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-232">Type</span></span>| <span data-ttu-id="d4854-233">描述</span><span class="sxs-lookup"><span data-stu-id="d4854-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d4854-234">字符串</span><span class="sxs-lookup"><span data-stu-id="d4854-234">String</span></span>|<span data-ttu-id="d4854-235">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="d4854-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d4854-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d4854-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="d4854-237">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="d4854-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-238">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-238">Requirements</span></span>

|<span data-ttu-id="d4854-239">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-239">Requirement</span></span>| <span data-ttu-id="d4854-240">值</span><span class="sxs-lookup"><span data-stu-id="d4854-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-242">1.3</span><span class="sxs-lookup"><span data-stu-id="d4854-242">1.3</span></span>|
|[<span data-ttu-id="d4854-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-244">受限</span><span class="sxs-lookup"><span data-stu-id="d4854-244">Restricted</span></span>|
|[<span data-ttu-id="d4854-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d4854-247">返回：</span><span class="sxs-lookup"><span data-stu-id="d4854-247">Returns:</span></span>

<span data-ttu-id="d4854-248">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="d4854-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d4854-249">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-249">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="d4854-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="d4854-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="d4854-251">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="d4854-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d4854-p108">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="d4854-p108">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d4854-p109">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="d4854-p109">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-257">Parameters</span><span class="sxs-lookup"><span data-stu-id="d4854-257">Parameters</span></span>

|<span data-ttu-id="d4854-258">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-258">Name</span></span>| <span data-ttu-id="d4854-259">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-259">Type</span></span>| <span data-ttu-id="d4854-260">描述</span><span class="sxs-lookup"><span data-stu-id="d4854-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d4854-261">日期</span><span class="sxs-lookup"><span data-stu-id="d4854-261">Date</span></span>|<span data-ttu-id="d4854-262">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="d4854-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-263">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-263">Requirements</span></span>

|<span data-ttu-id="d4854-264">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-264">Requirement</span></span>| <span data-ttu-id="d4854-265">值</span><span class="sxs-lookup"><span data-stu-id="d4854-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-266">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-267">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-267">1.0</span></span>|
|[<span data-ttu-id="d4854-268">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-269">ReadItem</span></span>|
|[<span data-ttu-id="d4854-270">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-271">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d4854-272">返回：</span><span class="sxs-lookup"><span data-stu-id="d4854-272">Returns:</span></span>

<span data-ttu-id="d4854-273">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d4854-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d4854-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d4854-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d4854-275">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="d4854-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-276">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d4854-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d4854-p110">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="d4854-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-279">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-279">Parameters</span></span>

|<span data-ttu-id="d4854-280">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-280">Name</span></span>| <span data-ttu-id="d4854-281">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-281">Type</span></span>| <span data-ttu-id="d4854-282">描述</span><span class="sxs-lookup"><span data-stu-id="d4854-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d4854-283">字符串</span><span class="sxs-lookup"><span data-stu-id="d4854-283">String</span></span>|<span data-ttu-id="d4854-284">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="d4854-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d4854-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d4854-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="d4854-286">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="d4854-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-287">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-287">Requirements</span></span>

|<span data-ttu-id="d4854-288">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-288">Requirement</span></span>| <span data-ttu-id="d4854-289">值</span><span class="sxs-lookup"><span data-stu-id="d4854-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-290">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-291">1.3</span><span class="sxs-lookup"><span data-stu-id="d4854-291">1.3</span></span>|
|[<span data-ttu-id="d4854-292">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-293">受限</span><span class="sxs-lookup"><span data-stu-id="d4854-293">Restricted</span></span>|
|[<span data-ttu-id="d4854-294">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-295">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d4854-296">返回：</span><span class="sxs-lookup"><span data-stu-id="d4854-296">Returns:</span></span>

<span data-ttu-id="d4854-297">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="d4854-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d4854-298">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-298">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d4854-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d4854-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d4854-300">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="d4854-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d4854-301">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="d4854-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-302">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-302">Parameters</span></span>

|<span data-ttu-id="d4854-303">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-303">Name</span></span>| <span data-ttu-id="d4854-304">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-304">Type</span></span>| <span data-ttu-id="d4854-305">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d4854-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d4854-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="d4854-307">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="d4854-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-308">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-308">Requirements</span></span>

|<span data-ttu-id="d4854-309">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-309">Requirement</span></span>| <span data-ttu-id="d4854-310">值</span><span class="sxs-lookup"><span data-stu-id="d4854-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-312">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-312">1.0</span></span>|
|[<span data-ttu-id="d4854-313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-314">ReadItem</span></span>|
|[<span data-ttu-id="d4854-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-316">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d4854-317">返回：</span><span class="sxs-lookup"><span data-stu-id="d4854-317">Returns:</span></span>

<span data-ttu-id="d4854-318">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="d4854-318">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="d4854-319">键入：日期</span><span class="sxs-lookup"><span data-stu-id="d4854-319">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="d4854-320">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-320">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="d4854-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d4854-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d4854-322">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="d4854-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-323">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d4854-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d4854-324">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="d4854-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d4854-p111">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="d4854-p111">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d4854-327">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="d4854-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d4854-328">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="d4854-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-329">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-329">Parameters</span></span>

|<span data-ttu-id="d4854-330">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-330">Name</span></span>| <span data-ttu-id="d4854-331">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-331">Type</span></span>| <span data-ttu-id="d4854-332">描述</span><span class="sxs-lookup"><span data-stu-id="d4854-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d4854-333">字符串</span><span class="sxs-lookup"><span data-stu-id="d4854-333">String</span></span>|<span data-ttu-id="d4854-334">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="d4854-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-335">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-335">Requirements</span></span>

|<span data-ttu-id="d4854-336">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-336">Requirement</span></span>| <span data-ttu-id="d4854-337">值</span><span class="sxs-lookup"><span data-stu-id="d4854-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-338">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-339">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-339">1.0</span></span>|
|[<span data-ttu-id="d4854-340">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-341">ReadItem</span></span>|
|[<span data-ttu-id="d4854-342">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-343">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4854-344">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="d4854-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d4854-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d4854-346">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="d4854-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-347">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d4854-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d4854-348">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="d4854-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d4854-349">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="d4854-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d4854-350">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="d4854-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d4854-p112">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="d4854-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-353">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-353">Parameters</span></span>

|<span data-ttu-id="d4854-354">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-354">Name</span></span>| <span data-ttu-id="d4854-355">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-355">Type</span></span>| <span data-ttu-id="d4854-356">描述</span><span class="sxs-lookup"><span data-stu-id="d4854-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d4854-357">字符串</span><span class="sxs-lookup"><span data-stu-id="d4854-357">String</span></span>|<span data-ttu-id="d4854-358">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="d4854-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-359">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-359">Requirements</span></span>

|<span data-ttu-id="d4854-360">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-360">Requirement</span></span>| <span data-ttu-id="d4854-361">值</span><span class="sxs-lookup"><span data-stu-id="d4854-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-363">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-363">1.0</span></span>|
|[<span data-ttu-id="d4854-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-365">ReadItem</span></span>|
|[<span data-ttu-id="d4854-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4854-368">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d4854-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d4854-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d4854-370">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="d4854-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-371">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d4854-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d4854-p113">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="d4854-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d4854-p114">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="d4854-p114">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d4854-p115">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="d4854-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d4854-379">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="d4854-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-380">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-380">Parameters</span></span>

|<span data-ttu-id="d4854-381">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-381">Name</span></span>| <span data-ttu-id="d4854-382">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-382">Type</span></span>| <span data-ttu-id="d4854-383">描述</span><span class="sxs-lookup"><span data-stu-id="d4854-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d4854-384">对象</span><span class="sxs-lookup"><span data-stu-id="d4854-384">Object</span></span> | <span data-ttu-id="d4854-385">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="d4854-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d4854-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="d4854-p116">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="d4854-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d4854-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="d4854-p117">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="d4854-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d4854-392">日期</span><span class="sxs-lookup"><span data-stu-id="d4854-392">Date</span></span> | <span data-ttu-id="d4854-393">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="d4854-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d4854-394">Date</span><span class="sxs-lookup"><span data-stu-id="d4854-394">Date</span></span> | <span data-ttu-id="d4854-395">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="d4854-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d4854-396">String</span><span class="sxs-lookup"><span data-stu-id="d4854-396">String</span></span> | <span data-ttu-id="d4854-p118">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d4854-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d4854-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d4854-p119">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="d4854-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d4854-402">String</span><span class="sxs-lookup"><span data-stu-id="d4854-402">String</span></span> | <span data-ttu-id="d4854-p120">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d4854-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d4854-405">字符串</span><span class="sxs-lookup"><span data-stu-id="d4854-405">String</span></span> | <span data-ttu-id="d4854-p121">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d4854-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d4854-408">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-408">Requirements</span></span>

|<span data-ttu-id="d4854-409">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-409">Requirement</span></span>| <span data-ttu-id="d4854-410">值</span><span class="sxs-lookup"><span data-stu-id="d4854-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-412">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-412">1.0</span></span>|
|[<span data-ttu-id="d4854-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-414">ReadItem</span></span>|
|[<span data-ttu-id="d4854-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-416">阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4854-417">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d4854-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d4854-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d4854-419">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="d4854-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d4854-p122">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="d4854-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-422">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="d4854-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="d4854-423">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="d4854-423">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="d4854-424">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="d4854-424">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="d4854-425">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="d4854-425">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="d4854-426">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="d4854-426">**REST Tokens**</span></span>

<span data-ttu-id="d4854-p124">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="d4854-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d4854-430">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="d4854-430">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d4854-431">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="d4854-431">**EWS Tokens**</span></span>

<span data-ttu-id="d4854-p125">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="d4854-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d4854-434">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="d4854-434">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="d4854-435">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="d4854-435">You can pass the token and an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="d4854-436">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="d4854-436">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="d4854-437">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="d4854-437">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-438">Parameters</span><span class="sxs-lookup"><span data-stu-id="d4854-438">Parameters</span></span>

|<span data-ttu-id="d4854-439">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-439">Name</span></span>| <span data-ttu-id="d4854-440">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-440">Type</span></span>| <span data-ttu-id="d4854-441">属性</span><span class="sxs-lookup"><span data-stu-id="d4854-441">Attributes</span></span>| <span data-ttu-id="d4854-442">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-442">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d4854-443">Object</span><span class="sxs-lookup"><span data-stu-id="d4854-443">Object</span></span> | <span data-ttu-id="d4854-444">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-444">&lt;optional&gt;</span></span> | <span data-ttu-id="d4854-445">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d4854-445">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d4854-446">布尔值</span><span class="sxs-lookup"><span data-stu-id="d4854-446">Boolean</span></span> |  <span data-ttu-id="d4854-447">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-447">&lt;optional&gt;</span></span> | <span data-ttu-id="d4854-p127">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="d4854-p127">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d4854-450">Object</span><span class="sxs-lookup"><span data-stu-id="d4854-450">Object</span></span> |  <span data-ttu-id="d4854-451">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-451">&lt;optional&gt;</span></span> | <span data-ttu-id="d4854-452">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="d4854-452">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d4854-453">函数</span><span class="sxs-lookup"><span data-stu-id="d4854-453">function</span></span>||<span data-ttu-id="d4854-454">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d4854-454">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d4854-455">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="d4854-455">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d4854-456">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="d4854-456">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d4854-457">错误</span><span class="sxs-lookup"><span data-stu-id="d4854-457">Errors</span></span>

|<span data-ttu-id="d4854-458">错误代码</span><span class="sxs-lookup"><span data-stu-id="d4854-458">Error code</span></span>|<span data-ttu-id="d4854-459">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-459">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d4854-460">请求失败。</span><span class="sxs-lookup"><span data-stu-id="d4854-460">The request has failed.</span></span> <span data-ttu-id="d4854-461">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="d4854-461">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d4854-462">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="d4854-462">The Exchange server returned an error.</span></span> <span data-ttu-id="d4854-463">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="d4854-463">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d4854-464">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="d4854-464">The user is no longer connected to the network.</span></span> <span data-ttu-id="d4854-465">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="d4854-465">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-466">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-466">Requirements</span></span>

|<span data-ttu-id="d4854-467">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-467">Requirement</span></span>| <span data-ttu-id="d4854-468">值</span><span class="sxs-lookup"><span data-stu-id="d4854-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-470">1.5</span><span class="sxs-lookup"><span data-stu-id="d4854-470">1.5</span></span> |
|[<span data-ttu-id="d4854-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-472">ReadItem</span></span>|
|[<span data-ttu-id="d4854-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-474">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-474">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4854-475">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-475">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d4854-476">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d4854-476">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d4854-477">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="d4854-477">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d4854-p131">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="d4854-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d4854-480">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="d4854-480">You can pass the token and an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="d4854-481">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="d4854-481">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="d4854-482">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="d4854-482">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d4854-483">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="d4854-483">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="d4854-484">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="d4854-484">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="d4854-485">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="d4854-485">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-486">Parameters</span><span class="sxs-lookup"><span data-stu-id="d4854-486">Parameters</span></span>

|<span data-ttu-id="d4854-487">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-487">Name</span></span>| <span data-ttu-id="d4854-488">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-488">Type</span></span>| <span data-ttu-id="d4854-489">属性</span><span class="sxs-lookup"><span data-stu-id="d4854-489">Attributes</span></span>| <span data-ttu-id="d4854-490">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-490">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d4854-491">函数</span><span class="sxs-lookup"><span data-stu-id="d4854-491">function</span></span>||<span data-ttu-id="d4854-492">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d4854-492">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d4854-493">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="d4854-493">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d4854-494">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="d4854-494">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="d4854-495">对象</span><span class="sxs-lookup"><span data-stu-id="d4854-495">Object</span></span>| <span data-ttu-id="d4854-496">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-496">&lt;optional&gt;</span></span>|<span data-ttu-id="d4854-497">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="d4854-497">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d4854-498">错误</span><span class="sxs-lookup"><span data-stu-id="d4854-498">Errors</span></span>

|<span data-ttu-id="d4854-499">错误代码</span><span class="sxs-lookup"><span data-stu-id="d4854-499">Error code</span></span>|<span data-ttu-id="d4854-500">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-500">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d4854-501">请求失败。</span><span class="sxs-lookup"><span data-stu-id="d4854-501">The request has failed.</span></span> <span data-ttu-id="d4854-502">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="d4854-502">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d4854-503">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="d4854-503">The Exchange server returned an error.</span></span> <span data-ttu-id="d4854-504">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="d4854-504">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d4854-505">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="d4854-505">The user is no longer connected to the network.</span></span> <span data-ttu-id="d4854-506">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="d4854-506">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-507">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-507">Requirements</span></span>

|<span data-ttu-id="d4854-508">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-508">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d4854-509">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-510">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-510">1.0</span></span> | <span data-ttu-id="d4854-511">1.3</span><span class="sxs-lookup"><span data-stu-id="d4854-511">1.3</span></span> |
|[<span data-ttu-id="d4854-512">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-513">ReadItem</span></span> | <span data-ttu-id="d4854-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-514">ReadItem</span></span> |
|[<span data-ttu-id="d4854-515">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-515">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-516">阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-516">Read</span></span> | <span data-ttu-id="d4854-517">撰写</span><span class="sxs-lookup"><span data-stu-id="d4854-517">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="d4854-518">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-518">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d4854-519">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d4854-519">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d4854-520">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="d4854-520">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d4854-521">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="d4854-521">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-522">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-522">Parameters</span></span>

|<span data-ttu-id="d4854-523">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-523">Name</span></span>| <span data-ttu-id="d4854-524">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-524">Type</span></span>| <span data-ttu-id="d4854-525">属性</span><span class="sxs-lookup"><span data-stu-id="d4854-525">Attributes</span></span>| <span data-ttu-id="d4854-526">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-526">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d4854-527">函数</span><span class="sxs-lookup"><span data-stu-id="d4854-527">function</span></span>||<span data-ttu-id="d4854-528">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d4854-528">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d4854-529">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="d4854-529">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d4854-530">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="d4854-530">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="d4854-531">对象</span><span class="sxs-lookup"><span data-stu-id="d4854-531">Object</span></span>| <span data-ttu-id="d4854-532">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-532">&lt;optional&gt;</span></span>|<span data-ttu-id="d4854-533">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="d4854-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d4854-534">错误</span><span class="sxs-lookup"><span data-stu-id="d4854-534">Errors</span></span>

|<span data-ttu-id="d4854-535">错误代码</span><span class="sxs-lookup"><span data-stu-id="d4854-535">Error code</span></span>|<span data-ttu-id="d4854-536">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-536">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d4854-537">请求失败。</span><span class="sxs-lookup"><span data-stu-id="d4854-537">The request has failed.</span></span> <span data-ttu-id="d4854-538">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="d4854-538">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d4854-539">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="d4854-539">The Exchange server returned an error.</span></span> <span data-ttu-id="d4854-540">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="d4854-540">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d4854-541">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="d4854-541">The user is no longer connected to the network.</span></span> <span data-ttu-id="d4854-542">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="d4854-542">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-543">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-543">Requirements</span></span>

|<span data-ttu-id="d4854-544">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-544">Requirement</span></span>| <span data-ttu-id="d4854-545">值</span><span class="sxs-lookup"><span data-stu-id="d4854-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-546">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-547">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-547">1.0</span></span>|
|[<span data-ttu-id="d4854-548">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-549">ReadItem</span></span>|
|[<span data-ttu-id="d4854-550">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-551">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4854-552">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-552">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d4854-553">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d4854-553">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d4854-554">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="d4854-554">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-555">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="d4854-555">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d4854-556">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="d4854-556">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="d4854-557">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="d4854-557">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d4854-558">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="d4854-558">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d4854-559">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="d4854-559">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="d4854-560">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="d4854-560">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d4854-561">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="d4854-561">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d4854-562">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="d4854-562">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d4854-p141">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="d4854-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d4854-565">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="d4854-565">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d4854-566">版本差异</span><span class="sxs-lookup"><span data-stu-id="d4854-566">Version differences</span></span>

<span data-ttu-id="d4854-567">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="d4854-567">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d4854-p142">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="d4854-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-571">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-571">Parameters</span></span>

|<span data-ttu-id="d4854-572">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-572">Name</span></span>| <span data-ttu-id="d4854-573">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-573">Type</span></span>| <span data-ttu-id="d4854-574">属性</span><span class="sxs-lookup"><span data-stu-id="d4854-574">Attributes</span></span>| <span data-ttu-id="d4854-575">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-575">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d4854-576">字符串</span><span class="sxs-lookup"><span data-stu-id="d4854-576">String</span></span>||<span data-ttu-id="d4854-577">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="d4854-577">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d4854-578">函数</span><span class="sxs-lookup"><span data-stu-id="d4854-578">function</span></span>||<span data-ttu-id="d4854-579">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d4854-579">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d4854-580">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="d4854-580">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="d4854-581">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="d4854-581">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="d4854-582">对象</span><span class="sxs-lookup"><span data-stu-id="d4854-582">Object</span></span>| <span data-ttu-id="d4854-583">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-583">&lt;optional&gt;</span></span>|<span data-ttu-id="d4854-584">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="d4854-584">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-585">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-585">Requirements</span></span>

|<span data-ttu-id="d4854-586">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-586">Requirement</span></span>| <span data-ttu-id="d4854-587">值</span><span class="sxs-lookup"><span data-stu-id="d4854-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-588">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-589">1.0</span><span class="sxs-lookup"><span data-stu-id="d4854-589">1.0</span></span>|
|[<span data-ttu-id="d4854-590">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-591">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d4854-591">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d4854-592">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-593">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-593">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4854-594">示例</span><span class="sxs-lookup"><span data-stu-id="d4854-594">Example</span></span>

<span data-ttu-id="d4854-595">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="d4854-595">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d4854-596">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d4854-596">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d4854-597">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d4854-597">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d4854-598">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="d4854-598">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4854-599">参数</span><span class="sxs-lookup"><span data-stu-id="d4854-599">Parameters</span></span>

| <span data-ttu-id="d4854-600">名称</span><span class="sxs-lookup"><span data-stu-id="d4854-600">Name</span></span> | <span data-ttu-id="d4854-601">类型</span><span class="sxs-lookup"><span data-stu-id="d4854-601">Type</span></span> | <span data-ttu-id="d4854-602">属性</span><span class="sxs-lookup"><span data-stu-id="d4854-602">Attributes</span></span> | <span data-ttu-id="d4854-603">说明</span><span class="sxs-lookup"><span data-stu-id="d4854-603">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d4854-604">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d4854-604">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d4854-605">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="d4854-605">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d4854-606">对象</span><span class="sxs-lookup"><span data-stu-id="d4854-606">Object</span></span> | <span data-ttu-id="d4854-607">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-607">&lt;optional&gt;</span></span> | <span data-ttu-id="d4854-608">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d4854-608">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d4854-609">对象</span><span class="sxs-lookup"><span data-stu-id="d4854-609">Object</span></span> | <span data-ttu-id="d4854-610">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-610">&lt;optional&gt;</span></span> | <span data-ttu-id="d4854-611">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d4854-611">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d4854-612">函数</span><span class="sxs-lookup"><span data-stu-id="d4854-612">function</span></span>| <span data-ttu-id="d4854-613">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d4854-613">&lt;optional&gt;</span></span>|<span data-ttu-id="d4854-614">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d4854-614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4854-615">Requirements</span><span class="sxs-lookup"><span data-stu-id="d4854-615">Requirements</span></span>

|<span data-ttu-id="d4854-616">要求</span><span class="sxs-lookup"><span data-stu-id="d4854-616">Requirement</span></span>| <span data-ttu-id="d4854-617">值</span><span class="sxs-lookup"><span data-stu-id="d4854-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4854-618">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d4854-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4854-619">1.5</span><span class="sxs-lookup"><span data-stu-id="d4854-619">1.5</span></span> |
|[<span data-ttu-id="d4854-620">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d4854-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4854-621">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4854-621">ReadItem</span></span> |
|[<span data-ttu-id="d4854-622">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d4854-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4854-623">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d4854-623">Compose or Read</span></span>|
