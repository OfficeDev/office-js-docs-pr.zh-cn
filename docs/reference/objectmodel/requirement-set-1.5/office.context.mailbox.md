---
title: Office.context.mailbox - 要求集 1.5
description: ''
ms.date: 04/24/2019
localization_priority: Priority
ms.openlocfilehash: a0c1a45fd3eaa9cf324a6854120d642eb7520132
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127280"
---
# <a name="mailbox"></a><span data-ttu-id="b1ead-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="b1ead-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="b1ead-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="b1ead-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="b1ead-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="b1ead-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ead-105">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-105">Requirements</span></span>

|<span data-ttu-id="b1ead-106">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-106">Requirement</span></span>| <span data-ttu-id="b1ead-107">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-109">1.0</span></span>|
|[<span data-ttu-id="b1ead-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-111">受限</span><span class="sxs-lookup"><span data-stu-id="b1ead-111">Restricted</span></span>|
|[<span data-ttu-id="b1ead-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b1ead-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-114">Members and methods</span></span>

| <span data-ttu-id="b1ead-115">成员</span><span class="sxs-lookup"><span data-stu-id="b1ead-115">Member</span></span> | <span data-ttu-id="b1ead-116">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b1ead-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="b1ead-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="b1ead-118">成员</span><span class="sxs-lookup"><span data-stu-id="b1ead-118">Member</span></span> |
| [<span data-ttu-id="b1ead-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="b1ead-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="b1ead-120">成员</span><span class="sxs-lookup"><span data-stu-id="b1ead-120">Member</span></span> |
| [<span data-ttu-id="b1ead-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b1ead-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="b1ead-122">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-122">Method</span></span> |
| [<span data-ttu-id="b1ead-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="b1ead-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="b1ead-124">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-124">Method</span></span> |
| [<span data-ttu-id="b1ead-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="b1ead-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="b1ead-126">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-126">Method</span></span> |
| [<span data-ttu-id="b1ead-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="b1ead-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="b1ead-128">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-128">Method</span></span> |
| [<span data-ttu-id="b1ead-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="b1ead-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="b1ead-130">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-130">Method</span></span> |
| [<span data-ttu-id="b1ead-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b1ead-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="b1ead-132">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-132">Method</span></span> |
| [<span data-ttu-id="b1ead-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="b1ead-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="b1ead-134">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-134">Method</span></span> |
| [<span data-ttu-id="b1ead-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b1ead-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="b1ead-136">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-136">Method</span></span> |
| [<span data-ttu-id="b1ead-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b1ead-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="b1ead-138">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-138">Method</span></span> |
| [<span data-ttu-id="b1ead-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b1ead-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="b1ead-140">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-140">Method</span></span> |
| [<span data-ttu-id="b1ead-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b1ead-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="b1ead-142">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-142">Method</span></span> |
| [<span data-ttu-id="b1ead-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="b1ead-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="b1ead-144">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-144">Method</span></span> |
| [<span data-ttu-id="b1ead-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b1ead-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="b1ead-146">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b1ead-147">命名空间</span><span class="sxs-lookup"><span data-stu-id="b1ead-147">Namespaces</span></span>

<span data-ttu-id="b1ead-148">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="b1ead-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="b1ead-149">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="b1ead-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="b1ead-150">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="b1ead-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="b1ead-151">Members</span><span class="sxs-lookup"><span data-stu-id="b1ead-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="b1ead-152">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="b1ead-152">ewsUrl :String</span></span>

<span data-ttu-id="b1ead-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-155">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="b1ead-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ead-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="b1ead-158">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="b1ead-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="b1ead-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ead-161">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-161">Type</span></span>

*   <span data-ttu-id="b1ead-162">String</span><span class="sxs-lookup"><span data-stu-id="b1ead-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ead-163">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-163">Requirements</span></span>

|<span data-ttu-id="b1ead-164">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-164">Requirement</span></span>| <span data-ttu-id="b1ead-165">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-166">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-167">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-167">1.0</span></span>|
|[<span data-ttu-id="b1ead-168">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-169">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-171">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="b1ead-172">restUrl：String</span><span class="sxs-lookup"><span data-stu-id="b1ead-172">restUrl :String</span></span>

<span data-ttu-id="b1ead-173">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="b1ead-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="b1ead-174">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="b1ead-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="b1ead-175">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="b1ead-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="b1ead-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-178">连接到配置了自定义 REST URL 的 Exchange 2016 或更高版本本地安装的 Outlook 客户端将返回 `restUrl` 的无效值。</span><span class="sxs-lookup"><span data-stu-id="b1ead-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="b1ead-179">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-179">Type</span></span>

*   <span data-ttu-id="b1ead-180">String</span><span class="sxs-lookup"><span data-stu-id="b1ead-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b1ead-181">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-181">Requirements</span></span>

|<span data-ttu-id="b1ead-182">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-182">Requirement</span></span>| <span data-ttu-id="b1ead-183">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-185">1.5</span><span class="sxs-lookup"><span data-stu-id="b1ead-185">1.5</span></span> |
|[<span data-ttu-id="b1ead-186">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-187">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="b1ead-190">方法</span><span class="sxs-lookup"><span data-stu-id="b1ead-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="b1ead-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b1ead-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="b1ead-192">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="b1ead-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="b1ead-193">目前，唯一受支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该类型。</span><span class="sxs-lookup"><span data-stu-id="b1ead-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="b1ead-194">此事件由实现可固定任务窗格的加载项使用，并允许加载项根据当前选定的项目刷新任务窗格 UI。</span><span class="sxs-lookup"><span data-stu-id="b1ead-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-195">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-195">Parameters</span></span>

| <span data-ttu-id="b1ead-196">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-196">Name</span></span> | <span data-ttu-id="b1ead-197">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-197">Type</span></span> | <span data-ttu-id="b1ead-198">属性</span><span class="sxs-lookup"><span data-stu-id="b1ead-198">Attributes</span></span> | <span data-ttu-id="b1ead-199">说明</span><span class="sxs-lookup"><span data-stu-id="b1ead-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b1ead-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b1ead-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b1ead-201">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="b1ead-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="b1ead-202">函数</span><span class="sxs-lookup"><span data-stu-id="b1ead-202">Function</span></span> || <span data-ttu-id="b1ead-p106">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="b1ead-206">Object</span><span class="sxs-lookup"><span data-stu-id="b1ead-206">Object</span></span> | <span data-ttu-id="b1ead-207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-207">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ead-208">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="b1ead-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b1ead-209">对象</span><span class="sxs-lookup"><span data-stu-id="b1ead-209">Object</span></span> | <span data-ttu-id="b1ead-210">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-210">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ead-211">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="b1ead-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b1ead-212">函数</span><span class="sxs-lookup"><span data-stu-id="b1ead-212">function</span></span>| <span data-ttu-id="b1ead-213">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-213">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ead-214">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b1ead-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-215">Requirements</span><span class="sxs-lookup"><span data-stu-id="b1ead-215">Requirements</span></span>

|<span data-ttu-id="b1ead-216">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-216">Requirement</span></span>| <span data-ttu-id="b1ead-217">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-219">1.5</span><span class="sxs-lookup"><span data-stu-id="b1ead-219">1.5</span></span> |
|[<span data-ttu-id="b1ead-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-221">ReadItem</span></span> |
|[<span data-ttu-id="b1ead-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-223">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ead-224">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="b1ead-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="b1ead-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="b1ead-226">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="b1ead-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-227">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b1ead-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ead-p107">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-230">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-230">Parameters</span></span>

|<span data-ttu-id="b1ead-231">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-231">Name</span></span>| <span data-ttu-id="b1ead-232">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-232">Type</span></span>| <span data-ttu-id="b1ead-233">描述</span><span class="sxs-lookup"><span data-stu-id="b1ead-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b1ead-234">字符串</span><span class="sxs-lookup"><span data-stu-id="b1ead-234">String</span></span>|<span data-ttu-id="b1ead-235">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="b1ead-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="b1ead-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="b1ead-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="b1ead-237">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="b1ead-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-238">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-238">Requirements</span></span>

|<span data-ttu-id="b1ead-239">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-239">Requirement</span></span>| <span data-ttu-id="b1ead-240">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-242">1.3</span><span class="sxs-lookup"><span data-stu-id="b1ead-242">1.3</span></span>|
|[<span data-ttu-id="b1ead-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-244">受限</span><span class="sxs-lookup"><span data-stu-id="b1ead-244">Restricted</span></span>|
|[<span data-ttu-id="b1ead-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ead-247">返回：</span><span class="sxs-lookup"><span data-stu-id="b1ead-247">Returns:</span></span>

<span data-ttu-id="b1ead-248">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="b1ead-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="b1ead-249">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="b1ead-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="b1ead-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="b1ead-251">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="b1ead-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="b1ead-p108">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="b1ead-p109">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-257">Parameters</span><span class="sxs-lookup"><span data-stu-id="b1ead-257">Parameters</span></span>

|<span data-ttu-id="b1ead-258">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-258">Name</span></span>| <span data-ttu-id="b1ead-259">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-259">Type</span></span>| <span data-ttu-id="b1ead-260">描述</span><span class="sxs-lookup"><span data-stu-id="b1ead-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="b1ead-261">日期</span><span class="sxs-lookup"><span data-stu-id="b1ead-261">Date</span></span>|<span data-ttu-id="b1ead-262">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="b1ead-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-263">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-263">Requirements</span></span>

|<span data-ttu-id="b1ead-264">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-264">Requirement</span></span>| <span data-ttu-id="b1ead-265">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-266">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-267">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-267">1.0</span></span>|
|[<span data-ttu-id="b1ead-268">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-269">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-270">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-271">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ead-272">返回：</span><span class="sxs-lookup"><span data-stu-id="b1ead-272">Returns:</span></span>

<span data-ttu-id="b1ead-273">类型：[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="b1ead-273">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="b1ead-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="b1ead-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="b1ead-275">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="b1ead-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-276">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b1ead-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ead-p110">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-279">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-279">Parameters</span></span>

|<span data-ttu-id="b1ead-280">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-280">Name</span></span>| <span data-ttu-id="b1ead-281">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-281">Type</span></span>| <span data-ttu-id="b1ead-282">描述</span><span class="sxs-lookup"><span data-stu-id="b1ead-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b1ead-283">字符串</span><span class="sxs-lookup"><span data-stu-id="b1ead-283">String</span></span>|<span data-ttu-id="b1ead-284">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="b1ead-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="b1ead-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="b1ead-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="b1ead-286">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="b1ead-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-287">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-287">Requirements</span></span>

|<span data-ttu-id="b1ead-288">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-288">Requirement</span></span>| <span data-ttu-id="b1ead-289">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-290">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-291">1.3</span><span class="sxs-lookup"><span data-stu-id="b1ead-291">1.3</span></span>|
|[<span data-ttu-id="b1ead-292">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-293">受限</span><span class="sxs-lookup"><span data-stu-id="b1ead-293">Restricted</span></span>|
|[<span data-ttu-id="b1ead-294">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-295">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ead-296">返回：</span><span class="sxs-lookup"><span data-stu-id="b1ead-296">Returns:</span></span>

<span data-ttu-id="b1ead-297">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="b1ead-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="b1ead-298">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="b1ead-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="b1ead-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="b1ead-300">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="b1ead-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="b1ead-301">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="b1ead-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-302">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-302">Parameters</span></span>

|<span data-ttu-id="b1ead-303">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-303">Name</span></span>| <span data-ttu-id="b1ead-304">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-304">Type</span></span>| <span data-ttu-id="b1ead-305">说明</span><span class="sxs-lookup"><span data-stu-id="b1ead-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="b1ead-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="b1ead-306">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="b1ead-307">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="b1ead-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-308">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-308">Requirements</span></span>

|<span data-ttu-id="b1ead-309">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-309">Requirement</span></span>| <span data-ttu-id="b1ead-310">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-312">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-312">1.0</span></span>|
|[<span data-ttu-id="b1ead-313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-314">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-316">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b1ead-317">返回：</span><span class="sxs-lookup"><span data-stu-id="b1ead-317">Returns:</span></span>

<span data-ttu-id="b1ead-318">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="b1ead-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="b1ead-319">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="b1ead-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b1ead-320">日期</span><span class="sxs-lookup"><span data-stu-id="b1ead-320">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="b1ead-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="b1ead-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="b1ead-322">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="b1ead-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-323">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b1ead-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ead-324">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="b1ead-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="b1ead-p111">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="b1ead-327">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="b1ead-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="b1ead-328">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="b1ead-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-329">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-329">Parameters</span></span>

|<span data-ttu-id="b1ead-330">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-330">Name</span></span>| <span data-ttu-id="b1ead-331">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-331">Type</span></span>| <span data-ttu-id="b1ead-332">描述</span><span class="sxs-lookup"><span data-stu-id="b1ead-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b1ead-333">字符串</span><span class="sxs-lookup"><span data-stu-id="b1ead-333">String</span></span>|<span data-ttu-id="b1ead-334">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="b1ead-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-335">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-335">Requirements</span></span>

|<span data-ttu-id="b1ead-336">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-336">Requirement</span></span>| <span data-ttu-id="b1ead-337">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-338">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-339">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-339">1.0</span></span>|
|[<span data-ttu-id="b1ead-340">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-341">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-342">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-343">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ead-344">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="b1ead-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="b1ead-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="b1ead-346">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="b1ead-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-347">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b1ead-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ead-348">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="b1ead-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="b1ead-349">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="b1ead-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="b1ead-350">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="b1ead-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="b1ead-p112">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-353">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-353">Parameters</span></span>

|<span data-ttu-id="b1ead-354">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-354">Name</span></span>| <span data-ttu-id="b1ead-355">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-355">Type</span></span>| <span data-ttu-id="b1ead-356">描述</span><span class="sxs-lookup"><span data-stu-id="b1ead-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b1ead-357">字符串</span><span class="sxs-lookup"><span data-stu-id="b1ead-357">String</span></span>|<span data-ttu-id="b1ead-358">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="b1ead-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-359">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-359">Requirements</span></span>

|<span data-ttu-id="b1ead-360">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-360">Requirement</span></span>| <span data-ttu-id="b1ead-361">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-363">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-363">1.0</span></span>|
|[<span data-ttu-id="b1ead-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-365">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ead-368">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="b1ead-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="b1ead-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="b1ead-370">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="b1ead-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-371">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b1ead-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b1ead-p113">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="b1ead-p114">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="b1ead-p115">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="b1ead-379">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="b1ead-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-380">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-380">Parameters</span></span>

|<span data-ttu-id="b1ead-381">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-381">Name</span></span>| <span data-ttu-id="b1ead-382">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-382">Type</span></span>| <span data-ttu-id="b1ead-383">描述</span><span class="sxs-lookup"><span data-stu-id="b1ead-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="b1ead-384">对象</span><span class="sxs-lookup"><span data-stu-id="b1ead-384">Object</span></span> | <span data-ttu-id="b1ead-385">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="b1ead-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="b1ead-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="b1ead-p116">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="b1ead-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="b1ead-p117">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="b1ead-392">日期</span><span class="sxs-lookup"><span data-stu-id="b1ead-392">Date</span></span> | <span data-ttu-id="b1ead-393">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="b1ead-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="b1ead-394">Date</span><span class="sxs-lookup"><span data-stu-id="b1ead-394">Date</span></span> | <span data-ttu-id="b1ead-395">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="b1ead-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="b1ead-396">String</span><span class="sxs-lookup"><span data-stu-id="b1ead-396">String</span></span> | <span data-ttu-id="b1ead-p118">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="b1ead-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="b1ead-p119">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="b1ead-402">String</span><span class="sxs-lookup"><span data-stu-id="b1ead-402">String</span></span> | <span data-ttu-id="b1ead-p120">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="b1ead-405">字符串</span><span class="sxs-lookup"><span data-stu-id="b1ead-405">String</span></span> | <span data-ttu-id="b1ead-p121">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b1ead-408">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-408">Requirements</span></span>

|<span data-ttu-id="b1ead-409">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-409">Requirement</span></span>| <span data-ttu-id="b1ead-410">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-412">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-412">1.0</span></span>|
|[<span data-ttu-id="b1ead-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-414">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-416">阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ead-417">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="b1ead-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b1ead-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="b1ead-419">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="b1ead-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="b1ead-p122">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-422">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="b1ead-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="b1ead-423">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="b1ead-423">**REST Tokens**</span></span>

<span data-ttu-id="b1ead-p123">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="b1ead-427">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="b1ead-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="b1ead-428">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="b1ead-428">**EWS Tokens**</span></span>

<span data-ttu-id="b1ead-p124">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="b1ead-431">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="b1ead-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-432">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-432">Parameters</span></span>

|<span data-ttu-id="b1ead-433">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-433">Name</span></span>| <span data-ttu-id="b1ead-434">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-434">Type</span></span>| <span data-ttu-id="b1ead-435">属性</span><span class="sxs-lookup"><span data-stu-id="b1ead-435">Attributes</span></span>| <span data-ttu-id="b1ead-436">说明</span><span class="sxs-lookup"><span data-stu-id="b1ead-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="b1ead-437">Object</span><span class="sxs-lookup"><span data-stu-id="b1ead-437">Object</span></span> | <span data-ttu-id="b1ead-438">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-438">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ead-439">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="b1ead-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="b1ead-440">布尔值</span><span class="sxs-lookup"><span data-stu-id="b1ead-440">Boolean</span></span> |  <span data-ttu-id="b1ead-441">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-441">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ead-p125">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p125">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b1ead-444">Object</span><span class="sxs-lookup"><span data-stu-id="b1ead-444">Object</span></span> |  <span data-ttu-id="b1ead-445">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-445">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ead-446">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="b1ead-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="b1ead-447">函数</span><span class="sxs-lookup"><span data-stu-id="b1ead-447">function</span></span>||<span data-ttu-id="b1ead-p126">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p126">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-450">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-450">Requirements</span></span>

|<span data-ttu-id="b1ead-451">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-451">Requirement</span></span>| <span data-ttu-id="b1ead-452">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-453">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-454">1.5</span><span class="sxs-lookup"><span data-stu-id="b1ead-454">1.5</span></span> |
|[<span data-ttu-id="b1ead-455">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-456">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-457">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-458">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-458">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ead-459">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-459">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="b1ead-460">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b1ead-460">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="b1ead-461">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="b1ead-461">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="b1ead-p127">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p127">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="b1ead-p128">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p128">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="b1ead-467">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="b1ead-467">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="b1ead-p129">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p129">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-470">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-470">Parameters</span></span>

|<span data-ttu-id="b1ead-471">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-471">Name</span></span>| <span data-ttu-id="b1ead-472">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-472">Type</span></span>| <span data-ttu-id="b1ead-473">属性</span><span class="sxs-lookup"><span data-stu-id="b1ead-473">Attributes</span></span>| <span data-ttu-id="b1ead-474">说明</span><span class="sxs-lookup"><span data-stu-id="b1ead-474">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b1ead-475">函数</span><span class="sxs-lookup"><span data-stu-id="b1ead-475">function</span></span>||<span data-ttu-id="b1ead-p130">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p130">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="b1ead-478">对象</span><span class="sxs-lookup"><span data-stu-id="b1ead-478">Object</span></span>| <span data-ttu-id="b1ead-479">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-479">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ead-480">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="b1ead-480">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-481">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-481">Requirements</span></span>

|<span data-ttu-id="b1ead-482">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-482">Requirement</span></span>| <span data-ttu-id="b1ead-483">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-485">1.3</span><span class="sxs-lookup"><span data-stu-id="b1ead-485">1.3</span></span>|
|[<span data-ttu-id="b1ead-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-487">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-489">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-489">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ead-490">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-490">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="b1ead-491">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b1ead-491">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="b1ead-492">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="b1ead-492">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="b1ead-493">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="b1ead-493">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-494">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-494">Parameters</span></span>

|<span data-ttu-id="b1ead-495">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-495">Name</span></span>| <span data-ttu-id="b1ead-496">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-496">Type</span></span>| <span data-ttu-id="b1ead-497">属性</span><span class="sxs-lookup"><span data-stu-id="b1ead-497">Attributes</span></span>| <span data-ttu-id="b1ead-498">说明</span><span class="sxs-lookup"><span data-stu-id="b1ead-498">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b1ead-499">函数</span><span class="sxs-lookup"><span data-stu-id="b1ead-499">function</span></span>||<span data-ttu-id="b1ead-500">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b1ead-500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b1ead-501">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="b1ead-501">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="b1ead-502">对象</span><span class="sxs-lookup"><span data-stu-id="b1ead-502">Object</span></span>| <span data-ttu-id="b1ead-503">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-503">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ead-504">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="b1ead-504">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-505">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-505">Requirements</span></span>

|<span data-ttu-id="b1ead-506">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-506">Requirement</span></span>| <span data-ttu-id="b1ead-507">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-509">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-509">1.0</span></span>|
|[<span data-ttu-id="b1ead-510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-511">ReadItem</span></span>|
|[<span data-ttu-id="b1ead-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-513">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-513">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ead-514">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-514">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="b1ead-515">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b1ead-515">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="b1ead-516">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="b1ead-516">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-517">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="b1ead-517">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="b1ead-518">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="b1ead-518">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="b1ead-519">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="b1ead-519">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="b1ead-520">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="b1ead-520">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="b1ead-521">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="b1ead-521">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="b1ead-522">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="b1ead-522">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="b1ead-523">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="b1ead-523">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="b1ead-524">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="b1ead-524">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="b1ead-p132">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p132">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="b1ead-527">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="b1ead-527">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="b1ead-528">版本差异</span><span class="sxs-lookup"><span data-stu-id="b1ead-528">Version differences</span></span>

<span data-ttu-id="b1ead-529">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="b1ead-529">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="b1ead-p133">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="b1ead-p133">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-533">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-533">Parameters</span></span>

|<span data-ttu-id="b1ead-534">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-534">Name</span></span>| <span data-ttu-id="b1ead-535">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-535">Type</span></span>| <span data-ttu-id="b1ead-536">属性</span><span class="sxs-lookup"><span data-stu-id="b1ead-536">Attributes</span></span>| <span data-ttu-id="b1ead-537">说明</span><span class="sxs-lookup"><span data-stu-id="b1ead-537">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b1ead-538">字符串</span><span class="sxs-lookup"><span data-stu-id="b1ead-538">String</span></span>||<span data-ttu-id="b1ead-539">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="b1ead-539">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="b1ead-540">函数</span><span class="sxs-lookup"><span data-stu-id="b1ead-540">function</span></span>||<span data-ttu-id="b1ead-541">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b1ead-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b1ead-542">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="b1ead-542">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="b1ead-543">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="b1ead-543">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="b1ead-544">对象</span><span class="sxs-lookup"><span data-stu-id="b1ead-544">Object</span></span>| <span data-ttu-id="b1ead-545">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-545">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ead-546">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="b1ead-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-547">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-547">Requirements</span></span>

|<span data-ttu-id="b1ead-548">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-548">Requirement</span></span>| <span data-ttu-id="b1ead-549">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-550">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-551">1.0</span><span class="sxs-lookup"><span data-stu-id="b1ead-551">1.0</span></span>|
|[<span data-ttu-id="b1ead-552">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-552">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-553">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="b1ead-553">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="b1ead-554">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-554">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-555">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-555">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b1ead-556">示例</span><span class="sxs-lookup"><span data-stu-id="b1ead-556">Example</span></span>

<span data-ttu-id="b1ead-557">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="b1ead-557">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="b1ead-558">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b1ead-558">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="b1ead-559">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="b1ead-559">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="b1ead-560">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="b1ead-560">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b1ead-561">参数</span><span class="sxs-lookup"><span data-stu-id="b1ead-561">Parameters</span></span>

| <span data-ttu-id="b1ead-562">名称</span><span class="sxs-lookup"><span data-stu-id="b1ead-562">Name</span></span> | <span data-ttu-id="b1ead-563">类型</span><span class="sxs-lookup"><span data-stu-id="b1ead-563">Type</span></span> | <span data-ttu-id="b1ead-564">属性</span><span class="sxs-lookup"><span data-stu-id="b1ead-564">Attributes</span></span> | <span data-ttu-id="b1ead-565">说明</span><span class="sxs-lookup"><span data-stu-id="b1ead-565">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b1ead-566">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b1ead-566">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b1ead-567">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="b1ead-567">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="b1ead-568">对象</span><span class="sxs-lookup"><span data-stu-id="b1ead-568">Object</span></span> | <span data-ttu-id="b1ead-569">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-569">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ead-570">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="b1ead-570">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b1ead-571">对象</span><span class="sxs-lookup"><span data-stu-id="b1ead-571">Object</span></span> | <span data-ttu-id="b1ead-572">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-572">&lt;optional&gt;</span></span> | <span data-ttu-id="b1ead-573">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="b1ead-573">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b1ead-574">函数</span><span class="sxs-lookup"><span data-stu-id="b1ead-574">function</span></span>| <span data-ttu-id="b1ead-575">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b1ead-575">&lt;optional&gt;</span></span>|<span data-ttu-id="b1ead-576">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b1ead-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b1ead-577">Requirements</span><span class="sxs-lookup"><span data-stu-id="b1ead-577">Requirements</span></span>

|<span data-ttu-id="b1ead-578">要求</span><span class="sxs-lookup"><span data-stu-id="b1ead-578">Requirement</span></span>| <span data-ttu-id="b1ead-579">值</span><span class="sxs-lookup"><span data-stu-id="b1ead-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="b1ead-580">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b1ead-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b1ead-581">1.5</span><span class="sxs-lookup"><span data-stu-id="b1ead-581">1.5</span></span> |
|[<span data-ttu-id="b1ead-582">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b1ead-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b1ead-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b1ead-583">ReadItem</span></span> |
|[<span data-ttu-id="b1ead-584">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b1ead-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b1ead-585">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b1ead-585">Compose or Read</span></span>|
