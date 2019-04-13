---
title: "\"context.subname\"-\"邮箱-要求集 1.7\""
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 0f84e657644b198fbca722a0628a5bafcce84377
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838548"
---
# <a name="mailbox"></a><span data-ttu-id="3b39e-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="3b39e-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="3b39e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="3b39e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="3b39e-104">为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3b39e-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3b39e-105">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-105">Requirements</span></span>

|<span data-ttu-id="3b39e-106">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-106">Requirement</span></span>| <span data-ttu-id="3b39e-107">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-109">1.0</span></span>|
|[<span data-ttu-id="3b39e-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-111">受限</span><span class="sxs-lookup"><span data-stu-id="3b39e-111">Restricted</span></span>|
|[<span data-ttu-id="3b39e-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3b39e-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-114">Members and methods</span></span>

| <span data-ttu-id="3b39e-115">成员</span><span class="sxs-lookup"><span data-stu-id="3b39e-115">Member</span></span> | <span data-ttu-id="3b39e-116">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3b39e-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="3b39e-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="3b39e-118">成员</span><span class="sxs-lookup"><span data-stu-id="3b39e-118">Member</span></span> |
| [<span data-ttu-id="3b39e-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="3b39e-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="3b39e-120">成员</span><span class="sxs-lookup"><span data-stu-id="3b39e-120">Member</span></span> |
| [<span data-ttu-id="3b39e-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3b39e-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="3b39e-122">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-122">Method</span></span> |
| [<span data-ttu-id="3b39e-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="3b39e-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="3b39e-124">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-124">Method</span></span> |
| [<span data-ttu-id="3b39e-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3b39e-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="3b39e-126">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-126">Method</span></span> |
| [<span data-ttu-id="3b39e-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="3b39e-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="3b39e-128">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-128">Method</span></span> |
| [<span data-ttu-id="3b39e-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="3b39e-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="3b39e-130">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-130">Method</span></span> |
| [<span data-ttu-id="3b39e-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3b39e-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="3b39e-132">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-132">Method</span></span> |
| [<span data-ttu-id="3b39e-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="3b39e-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="3b39e-134">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-134">Method</span></span> |
| [<span data-ttu-id="3b39e-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3b39e-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="3b39e-136">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-136">Method</span></span> |
| [<span data-ttu-id="3b39e-137">office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="3b39e-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="3b39e-138">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-138">Method</span></span> |
| [<span data-ttu-id="3b39e-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3b39e-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="3b39e-140">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-140">Method</span></span> |
| [<span data-ttu-id="3b39e-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3b39e-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="3b39e-142">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-142">Method</span></span> |
| [<span data-ttu-id="3b39e-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3b39e-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="3b39e-144">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-144">Method</span></span> |
| [<span data-ttu-id="3b39e-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="3b39e-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="3b39e-146">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-146">Method</span></span> |
| [<span data-ttu-id="3b39e-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3b39e-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="3b39e-148">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3b39e-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="3b39e-149">Namespaces</span></span>

<span data-ttu-id="3b39e-150">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="3b39e-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="3b39e-151">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="3b39e-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="3b39e-152">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="3b39e-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="3b39e-153">成员</span><span class="sxs-lookup"><span data-stu-id="3b39e-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="3b39e-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="3b39e-154">ewsUrl :String</span></span>

<span data-ttu-id="3b39e-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-157">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="3b39e-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3b39e-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3b39e-160">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="3b39e-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="3b39e-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3b39e-163">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-163">Type</span></span>

*   <span data-ttu-id="3b39e-164">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3b39e-165">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-165">Requirements</span></span>

|<span data-ttu-id="3b39e-166">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-166">Requirement</span></span>| <span data-ttu-id="3b39e-167">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-169">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-169">1.0</span></span>|
|[<span data-ttu-id="3b39e-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-171">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-173">Compose or Read</span></span>|

---
---

#### <a name="resturl-string"></a><span data-ttu-id="3b39e-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="3b39e-174">restUrl :String</span></span>

<span data-ttu-id="3b39e-175">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="3b39e-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="3b39e-176">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="3b39e-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="3b39e-177">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="3b39e-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="3b39e-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3b39e-180">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-180">Type</span></span>

*   <span data-ttu-id="3b39e-181">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3b39e-182">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-182">Requirements</span></span>

|<span data-ttu-id="3b39e-183">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-183">Requirement</span></span>| <span data-ttu-id="3b39e-184">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-185">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-186">1.5</span><span class="sxs-lookup"><span data-stu-id="3b39e-186">1.5</span></span> |
|[<span data-ttu-id="3b39e-187">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-188">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-189">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-190">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="3b39e-191">方法</span><span class="sxs-lookup"><span data-stu-id="3b39e-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="3b39e-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3b39e-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="3b39e-193">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="3b39e-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="3b39e-194">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="3b39e-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-195">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-195">Parameters</span></span>

| <span data-ttu-id="3b39e-196">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-196">Name</span></span> | <span data-ttu-id="3b39e-197">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-197">Type</span></span> | <span data-ttu-id="3b39e-198">属性</span><span class="sxs-lookup"><span data-stu-id="3b39e-198">Attributes</span></span> | <span data-ttu-id="3b39e-199">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3b39e-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3b39e-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3b39e-201">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="3b39e-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="3b39e-202">函数</span><span class="sxs-lookup"><span data-stu-id="3b39e-202">Function</span></span> || <span data-ttu-id="3b39e-p105">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="3b39e-206">Object</span><span class="sxs-lookup"><span data-stu-id="3b39e-206">Object</span></span> | <span data-ttu-id="3b39e-207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-207">&lt;optional&gt;</span></span> | <span data-ttu-id="3b39e-208">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3b39e-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3b39e-209">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-209">Object</span></span> | <span data-ttu-id="3b39e-210">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-210">&lt;optional&gt;</span></span> | <span data-ttu-id="3b39e-211">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3b39e-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3b39e-212">函数</span><span class="sxs-lookup"><span data-stu-id="3b39e-212">function</span></span>| <span data-ttu-id="3b39e-213">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-213">&lt;optional&gt;</span></span>|<span data-ttu-id="3b39e-214">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3b39e-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-215">Requirements</span><span class="sxs-lookup"><span data-stu-id="3b39e-215">Requirements</span></span>

|<span data-ttu-id="3b39e-216">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-216">Requirement</span></span>| <span data-ttu-id="3b39e-217">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-219">1.5</span><span class="sxs-lookup"><span data-stu-id="3b39e-219">1.5</span></span> |
|[<span data-ttu-id="3b39e-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-221">ReadItem</span></span> |
|[<span data-ttu-id="3b39e-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-223">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-224">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-224">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="3b39e-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3b39e-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3b39e-226">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="3b39e-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-227">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3b39e-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3b39e-p106">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-230">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-230">Parameters</span></span>

|<span data-ttu-id="3b39e-231">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-231">Name</span></span>| <span data-ttu-id="3b39e-232">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-232">Type</span></span>| <span data-ttu-id="3b39e-233">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3b39e-234">字符串</span><span class="sxs-lookup"><span data-stu-id="3b39e-234">String</span></span>|<span data-ttu-id="3b39e-235">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="3b39e-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="3b39e-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3b39e-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="3b39e-237">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="3b39e-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-238">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-238">Requirements</span></span>

|<span data-ttu-id="3b39e-239">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-239">Requirement</span></span>| <span data-ttu-id="3b39e-240">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-242">1.3</span><span class="sxs-lookup"><span data-stu-id="3b39e-242">1.3</span></span>|
|[<span data-ttu-id="3b39e-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-244">受限</span><span class="sxs-lookup"><span data-stu-id="3b39e-244">Restricted</span></span>|
|[<span data-ttu-id="3b39e-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3b39e-247">返回：</span><span class="sxs-lookup"><span data-stu-id="3b39e-247">Returns:</span></span>

<span data-ttu-id="3b39e-248">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="3b39e-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3b39e-249">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="3b39e-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="3b39e-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="3b39e-251">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="3b39e-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="3b39e-p107">Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="3b39e-p108">如果邮件应用程序在 Outlook 中运行，`convertToLocalClientTime` 方法将返回一个值设置为客户端计算机时区的字典对象。如果邮件应用程序在 Outlook Web App 中运行，`convertToLocalClientTime` 方法将返回值设置为 EAC 中指定的时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-257">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-257">Parameters</span></span>

|<span data-ttu-id="3b39e-258">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-258">Name</span></span>| <span data-ttu-id="3b39e-259">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-259">Type</span></span>| <span data-ttu-id="3b39e-260">描述</span><span class="sxs-lookup"><span data-stu-id="3b39e-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="3b39e-261">日期</span><span class="sxs-lookup"><span data-stu-id="3b39e-261">Date</span></span>|<span data-ttu-id="3b39e-262">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-263">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-263">Requirements</span></span>

|<span data-ttu-id="3b39e-264">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-264">Requirement</span></span>| <span data-ttu-id="3b39e-265">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-266">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-267">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-267">1.0</span></span>|
|[<span data-ttu-id="3b39e-268">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-269">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-270">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-271">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3b39e-272">返回：</span><span class="sxs-lookup"><span data-stu-id="3b39e-272">Returns:</span></span>

<span data-ttu-id="3b39e-273">类型：[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="3b39e-273">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

---
---

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="3b39e-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3b39e-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3b39e-275">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="3b39e-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-276">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3b39e-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3b39e-p109">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-279">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-279">Parameters</span></span>

|<span data-ttu-id="3b39e-280">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-280">Name</span></span>| <span data-ttu-id="3b39e-281">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-281">Type</span></span>| <span data-ttu-id="3b39e-282">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3b39e-283">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-283">String</span></span>|<span data-ttu-id="3b39e-284">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="3b39e-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="3b39e-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3b39e-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="3b39e-286">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="3b39e-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-287">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-287">Requirements</span></span>

|<span data-ttu-id="3b39e-288">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-288">Requirement</span></span>| <span data-ttu-id="3b39e-289">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-290">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-291">1.3</span><span class="sxs-lookup"><span data-stu-id="3b39e-291">1.3</span></span>|
|[<span data-ttu-id="3b39e-292">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-293">受限</span><span class="sxs-lookup"><span data-stu-id="3b39e-293">Restricted</span></span>|
|[<span data-ttu-id="3b39e-294">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-295">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3b39e-296">返回：</span><span class="sxs-lookup"><span data-stu-id="3b39e-296">Returns:</span></span>

<span data-ttu-id="3b39e-297">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="3b39e-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3b39e-298">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="3b39e-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="3b39e-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="3b39e-300">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3b39e-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="3b39e-301">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3b39e-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-302">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-302">Parameters</span></span>

|<span data-ttu-id="3b39e-303">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-303">Name</span></span>| <span data-ttu-id="3b39e-304">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-304">Type</span></span>| <span data-ttu-id="3b39e-305">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="3b39e-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3b39e-306">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="3b39e-307">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="3b39e-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-308">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-308">Requirements</span></span>

|<span data-ttu-id="3b39e-309">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-309">Requirement</span></span>| <span data-ttu-id="3b39e-310">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-312">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-312">1.0</span></span>|
|[<span data-ttu-id="3b39e-313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-314">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-316">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3b39e-317">返回：</span><span class="sxs-lookup"><span data-stu-id="3b39e-317">Returns:</span></span>

<span data-ttu-id="3b39e-318">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="3b39e-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="3b39e-319">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="3b39e-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3b39e-320">日期</span><span class="sxs-lookup"><span data-stu-id="3b39e-320">Date</span></span></dd>

</dl>

---
---

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="3b39e-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3b39e-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="3b39e-322">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="3b39e-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-323">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3b39e-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3b39e-324">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="3b39e-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3b39e-p110">在 Outlook for Mac 中，您可以使用此方法来显示不属于定期系列的单个约会，或显示定期系列的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="3b39e-327">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="3b39e-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="3b39e-328">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="3b39e-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-329">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-329">Parameters</span></span>

|<span data-ttu-id="3b39e-330">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-330">Name</span></span>| <span data-ttu-id="3b39e-331">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-331">Type</span></span>| <span data-ttu-id="3b39e-332">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3b39e-333">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-333">String</span></span>|<span data-ttu-id="3b39e-334">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="3b39e-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-335">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-335">Requirements</span></span>

|<span data-ttu-id="3b39e-336">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-336">Requirement</span></span>| <span data-ttu-id="3b39e-337">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-338">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-339">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-339">1.0</span></span>|
|[<span data-ttu-id="3b39e-340">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-341">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-342">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-343">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-344">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

####  <a name="displaymessageformitemid"></a><span data-ttu-id="3b39e-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3b39e-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="3b39e-346">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="3b39e-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-347">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3b39e-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3b39e-348">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="3b39e-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3b39e-349">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="3b39e-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="3b39e-350">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="3b39e-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="3b39e-p111">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-353">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-353">Parameters</span></span>

|<span data-ttu-id="3b39e-354">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-354">Name</span></span>| <span data-ttu-id="3b39e-355">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-355">Type</span></span>| <span data-ttu-id="3b39e-356">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3b39e-357">字符串</span><span class="sxs-lookup"><span data-stu-id="3b39e-357">String</span></span>|<span data-ttu-id="3b39e-358">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="3b39e-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-359">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-359">Requirements</span></span>

|<span data-ttu-id="3b39e-360">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-360">Requirement</span></span>| <span data-ttu-id="3b39e-361">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-363">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-363">1.0</span></span>|
|[<span data-ttu-id="3b39e-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-365">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-368">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="3b39e-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="3b39e-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="3b39e-370">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="3b39e-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-371">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3b39e-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3b39e-p112">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3b39e-p113">在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="3b39e-p114">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="3b39e-379">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="3b39e-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-380">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-381">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="3b39e-381">All parameters are optional.</span></span>

|<span data-ttu-id="3b39e-382">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-382">Name</span></span>| <span data-ttu-id="3b39e-383">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-383">Type</span></span>| <span data-ttu-id="3b39e-384">描述</span><span class="sxs-lookup"><span data-stu-id="3b39e-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3b39e-385">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-385">Object</span></span> | <span data-ttu-id="3b39e-386">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="3b39e-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="3b39e-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3b39e-p115">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="3b39e-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3b39e-p116">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="3b39e-393">日期</span><span class="sxs-lookup"><span data-stu-id="3b39e-393">Date</span></span> | <span data-ttu-id="3b39e-394">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3b39e-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="3b39e-395">Date</span><span class="sxs-lookup"><span data-stu-id="3b39e-395">Date</span></span> | <span data-ttu-id="3b39e-396">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3b39e-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="3b39e-397">字符串</span><span class="sxs-lookup"><span data-stu-id="3b39e-397">String</span></span> | <span data-ttu-id="3b39e-p117">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="3b39e-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="3b39e-p118">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3b39e-403">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-403">String</span></span> | <span data-ttu-id="3b39e-p119">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="3b39e-406">字符串</span><span class="sxs-lookup"><span data-stu-id="3b39e-406">String</span></span> | <span data-ttu-id="3b39e-p120">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3b39e-409">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-409">Requirements</span></span>

|<span data-ttu-id="3b39e-410">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-410">Requirement</span></span>| <span data-ttu-id="3b39e-411">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-412">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-413">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-413">1.0</span></span>|
|[<span data-ttu-id="3b39e-414">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-415">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-416">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-417">阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-418">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="3b39e-419">office.context.mailbox.displaynewmessageform (参数)</span><span class="sxs-lookup"><span data-stu-id="3b39e-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="3b39e-420">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="3b39e-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="3b39e-421">`displayNewMessageForm`方法将打开一个窗体, 使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="3b39e-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="3b39e-422">如果指定了参数, 则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="3b39e-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3b39e-423">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="3b39e-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-424">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-425">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="3b39e-425">All parameters are optional.</span></span>

|<span data-ttu-id="3b39e-426">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-426">Name</span></span>| <span data-ttu-id="3b39e-427">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-427">Type</span></span>| <span data-ttu-id="3b39e-428">描述</span><span class="sxs-lookup"><span data-stu-id="3b39e-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3b39e-429">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-429">Object</span></span> | <span data-ttu-id="3b39e-430">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="3b39e-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="3b39e-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3b39e-432">包含电子邮件地址的字符串数组, 或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3b39e-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="3b39e-433">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3b39e-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="3b39e-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3b39e-435">包含电子邮件地址的字符串数组, 或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3b39e-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="3b39e-436">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3b39e-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="3b39e-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="3b39e-438">包含电子邮件地址的字符串数组, 或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="3b39e-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="3b39e-439">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="3b39e-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3b39e-440">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-440">String</span></span> | <span data-ttu-id="3b39e-441">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="3b39e-441">A string containing the subject of the message.</span></span> <span data-ttu-id="3b39e-442">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3b39e-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="3b39e-443">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-443">String</span></span> | <span data-ttu-id="3b39e-444">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="3b39e-444">The HTML body of the message.</span></span> <span data-ttu-id="3b39e-445">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3b39e-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="3b39e-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3b39e-447">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="3b39e-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="3b39e-448">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-448">String</span></span> | <span data-ttu-id="3b39e-p127">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="3b39e-451">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-451">String</span></span> | <span data-ttu-id="3b39e-452">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3b39e-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="3b39e-453">String</span><span class="sxs-lookup"><span data-stu-id="3b39e-453">String</span></span> | <span data-ttu-id="3b39e-p128">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="3b39e-456">布尔</span><span class="sxs-lookup"><span data-stu-id="3b39e-456">Boolean</span></span> | <span data-ttu-id="3b39e-p129">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="3b39e-459">字符串</span><span class="sxs-lookup"><span data-stu-id="3b39e-459">String</span></span> | <span data-ttu-id="3b39e-460">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="3b39e-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="3b39e-461">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="3b39e-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="3b39e-462">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3b39e-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="3b39e-463">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-463">Requirements</span></span>

|<span data-ttu-id="3b39e-464">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-464">Requirement</span></span>| <span data-ttu-id="3b39e-465">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-466">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-467">1.6</span><span class="sxs-lookup"><span data-stu-id="3b39e-467">1.6</span></span> |
|[<span data-ttu-id="3b39e-468">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-469">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-470">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-471">阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-472">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-472">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="3b39e-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="3b39e-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="3b39e-474">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="3b39e-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="3b39e-p131">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-477">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="3b39e-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="3b39e-478">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="3b39e-478">**REST Tokens**</span></span>

<span data-ttu-id="3b39e-p132">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="3b39e-482">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="3b39e-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="3b39e-483">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="3b39e-483">**EWS Tokens**</span></span>

<span data-ttu-id="3b39e-p133">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="3b39e-486">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="3b39e-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-487">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-487">Parameters</span></span>

|<span data-ttu-id="3b39e-488">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-488">Name</span></span>| <span data-ttu-id="3b39e-489">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-489">Type</span></span>| <span data-ttu-id="3b39e-490">属性</span><span class="sxs-lookup"><span data-stu-id="3b39e-490">Attributes</span></span>| <span data-ttu-id="3b39e-491">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="3b39e-492">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-492">Object</span></span> | <span data-ttu-id="3b39e-493">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-493">&lt;optional&gt;</span></span> | <span data-ttu-id="3b39e-494">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3b39e-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="3b39e-495">布尔值</span><span class="sxs-lookup"><span data-stu-id="3b39e-495">Boolean</span></span> |  <span data-ttu-id="3b39e-496">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-496">&lt;optional&gt;</span></span> | <span data-ttu-id="3b39e-p134">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3b39e-499">Object</span><span class="sxs-lookup"><span data-stu-id="3b39e-499">Object</span></span> |  <span data-ttu-id="3b39e-500">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-500">&lt;optional&gt;</span></span> | <span data-ttu-id="3b39e-501">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3b39e-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="3b39e-502">函数</span><span class="sxs-lookup"><span data-stu-id="3b39e-502">function</span></span>||<span data-ttu-id="3b39e-p135">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-505">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-505">Requirements</span></span>

|<span data-ttu-id="3b39e-506">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-506">Requirement</span></span>| <span data-ttu-id="3b39e-507">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-509">1.5</span><span class="sxs-lookup"><span data-stu-id="3b39e-509">1.5</span></span> |
|[<span data-ttu-id="3b39e-510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-511">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-513">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-514">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="3b39e-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3b39e-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3b39e-516">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="3b39e-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="3b39e-p136">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="3b39e-p137">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3b39e-522">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="3b39e-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="3b39e-p138">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-525">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-525">Parameters</span></span>

|<span data-ttu-id="3b39e-526">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-526">Name</span></span>| <span data-ttu-id="3b39e-527">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-527">Type</span></span>| <span data-ttu-id="3b39e-528">属性</span><span class="sxs-lookup"><span data-stu-id="3b39e-528">Attributes</span></span>| <span data-ttu-id="3b39e-529">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3b39e-530">函数</span><span class="sxs-lookup"><span data-stu-id="3b39e-530">function</span></span>||<span data-ttu-id="3b39e-p139">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="3b39e-533">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-533">Object</span></span>| <span data-ttu-id="3b39e-534">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-534">&lt;optional&gt;</span></span>|<span data-ttu-id="3b39e-535">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3b39e-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-536">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-536">Requirements</span></span>

|<span data-ttu-id="3b39e-537">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-537">Requirement</span></span>| <span data-ttu-id="3b39e-538">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-539">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-540">1.3</span><span class="sxs-lookup"><span data-stu-id="3b39e-540">1.3</span></span>|
|[<span data-ttu-id="3b39e-541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-542">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-543">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-544">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-545">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-545">Example</span></span>

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

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="3b39e-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3b39e-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3b39e-547">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="3b39e-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="3b39e-548">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="3b39e-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-549">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-549">Parameters</span></span>

|<span data-ttu-id="3b39e-550">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-550">Name</span></span>| <span data-ttu-id="3b39e-551">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-551">Type</span></span>| <span data-ttu-id="3b39e-552">属性</span><span class="sxs-lookup"><span data-stu-id="3b39e-552">Attributes</span></span>| <span data-ttu-id="3b39e-553">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3b39e-554">function</span><span class="sxs-lookup"><span data-stu-id="3b39e-554">function</span></span>||<span data-ttu-id="3b39e-555">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3b39e-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3b39e-556">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3b39e-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="3b39e-557">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-557">Object</span></span>| <span data-ttu-id="3b39e-558">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-558">&lt;optional&gt;</span></span>|<span data-ttu-id="3b39e-559">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3b39e-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-560">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-560">Requirements</span></span>

|<span data-ttu-id="3b39e-561">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-561">Requirement</span></span>| <span data-ttu-id="3b39e-562">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-563">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-564">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-564">1.0</span></span>|
|[<span data-ttu-id="3b39e-565">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-566">ReadItem</span></span>|
|[<span data-ttu-id="3b39e-567">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-568">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-569">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-569">Example</span></span>

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

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="3b39e-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3b39e-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="3b39e-571">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="3b39e-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-572">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="3b39e-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="3b39e-573">在 Outlook for iOS 或 Outlook for Android 中</span><span class="sxs-lookup"><span data-stu-id="3b39e-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="3b39e-574">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="3b39e-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="3b39e-575">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="3b39e-575">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="3b39e-576">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="3b39e-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="3b39e-577">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="3b39e-577">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="3b39e-578">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="3b39e-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="3b39e-579">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="3b39e-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="3b39e-p141">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="3b39e-582">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="3b39e-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="3b39e-583">版本差异</span><span class="sxs-lookup"><span data-stu-id="3b39e-583">Version differences</span></span>

<span data-ttu-id="3b39e-584">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="3b39e-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="3b39e-p142">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="3b39e-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-588">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-588">Parameters</span></span>

|<span data-ttu-id="3b39e-589">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-589">Name</span></span>| <span data-ttu-id="3b39e-590">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-590">Type</span></span>| <span data-ttu-id="3b39e-591">属性</span><span class="sxs-lookup"><span data-stu-id="3b39e-591">Attributes</span></span>| <span data-ttu-id="3b39e-592">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="3b39e-593">字符串</span><span class="sxs-lookup"><span data-stu-id="3b39e-593">String</span></span>||<span data-ttu-id="3b39e-594">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="3b39e-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="3b39e-595">function</span><span class="sxs-lookup"><span data-stu-id="3b39e-595">function</span></span>||<span data-ttu-id="3b39e-596">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3b39e-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3b39e-597">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="3b39e-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="3b39e-598">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="3b39e-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="3b39e-599">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-599">Object</span></span>| <span data-ttu-id="3b39e-600">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-600">&lt;optional&gt;</span></span>|<span data-ttu-id="3b39e-601">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="3b39e-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-602">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-602">Requirements</span></span>

|<span data-ttu-id="3b39e-603">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-603">Requirement</span></span>| <span data-ttu-id="3b39e-604">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-605">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-606">1.0</span><span class="sxs-lookup"><span data-stu-id="3b39e-606">1.0</span></span>|
|[<span data-ttu-id="3b39e-607">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="3b39e-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="3b39e-609">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-610">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b39e-611">示例</span><span class="sxs-lookup"><span data-stu-id="3b39e-611">Example</span></span>

<span data-ttu-id="3b39e-612">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="3b39e-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="3b39e-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3b39e-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="3b39e-614">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="3b39e-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="3b39e-615">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="3b39e-615">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b39e-616">参数</span><span class="sxs-lookup"><span data-stu-id="3b39e-616">Parameters</span></span>

| <span data-ttu-id="3b39e-617">名称</span><span class="sxs-lookup"><span data-stu-id="3b39e-617">Name</span></span> | <span data-ttu-id="3b39e-618">类型</span><span class="sxs-lookup"><span data-stu-id="3b39e-618">Type</span></span> | <span data-ttu-id="3b39e-619">属性</span><span class="sxs-lookup"><span data-stu-id="3b39e-619">Attributes</span></span> | <span data-ttu-id="3b39e-620">说明</span><span class="sxs-lookup"><span data-stu-id="3b39e-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3b39e-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3b39e-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3b39e-622">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="3b39e-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="3b39e-623">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-623">Object</span></span> | <span data-ttu-id="3b39e-624">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-624">&lt;optional&gt;</span></span> | <span data-ttu-id="3b39e-625">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3b39e-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3b39e-626">对象</span><span class="sxs-lookup"><span data-stu-id="3b39e-626">Object</span></span> | <span data-ttu-id="3b39e-627">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-627">&lt;optional&gt;</span></span> | <span data-ttu-id="3b39e-628">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3b39e-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3b39e-629">函数</span><span class="sxs-lookup"><span data-stu-id="3b39e-629">function</span></span>| <span data-ttu-id="3b39e-630">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3b39e-630">&lt;optional&gt;</span></span>|<span data-ttu-id="3b39e-631">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3b39e-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b39e-632">Requirements</span><span class="sxs-lookup"><span data-stu-id="3b39e-632">Requirements</span></span>

|<span data-ttu-id="3b39e-633">要求</span><span class="sxs-lookup"><span data-stu-id="3b39e-633">Requirement</span></span>| <span data-ttu-id="3b39e-634">值</span><span class="sxs-lookup"><span data-stu-id="3b39e-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b39e-635">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3b39e-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b39e-636">1.5</span><span class="sxs-lookup"><span data-stu-id="3b39e-636">1.5</span></span> |
|[<span data-ttu-id="3b39e-637">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3b39e-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b39e-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b39e-638">ReadItem</span></span> |
|[<span data-ttu-id="3b39e-639">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3b39e-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b39e-640">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3b39e-640">Compose or Read</span></span>|
