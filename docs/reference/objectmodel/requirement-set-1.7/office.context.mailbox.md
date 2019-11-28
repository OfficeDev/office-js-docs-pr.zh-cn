---
title: "\"Context.subname\"-\"邮箱-要求集 1.7\""
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: c310ad38bb9821955fb0571d3693ce39715376f4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629670"
---
# <a name="mailbox"></a><span data-ttu-id="de52d-102">邮箱</span><span class="sxs-lookup"><span data-stu-id="de52d-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="de52d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="de52d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="de52d-104">为 Microsoft Outlook 提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="de52d-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="de52d-105">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-105">Requirements</span></span>

|<span data-ttu-id="de52d-106">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-106">Requirement</span></span>| <span data-ttu-id="de52d-107">值</span><span class="sxs-lookup"><span data-stu-id="de52d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-109">1.0</span></span>|
|[<span data-ttu-id="de52d-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-111">受限</span><span class="sxs-lookup"><span data-stu-id="de52d-111">Restricted</span></span>|
|[<span data-ttu-id="de52d-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="de52d-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="de52d-114">Members and methods</span></span>

| <span data-ttu-id="de52d-115">成员</span><span class="sxs-lookup"><span data-stu-id="de52d-115">Member</span></span> | <span data-ttu-id="de52d-116">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="de52d-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="de52d-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="de52d-118">成员</span><span class="sxs-lookup"><span data-stu-id="de52d-118">Member</span></span> |
| [<span data-ttu-id="de52d-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="de52d-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="de52d-120">成员</span><span class="sxs-lookup"><span data-stu-id="de52d-120">Member</span></span> |
| [<span data-ttu-id="de52d-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="de52d-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="de52d-122">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-122">Method</span></span> |
| [<span data-ttu-id="de52d-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="de52d-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="de52d-124">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-124">Method</span></span> |
| [<span data-ttu-id="de52d-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="de52d-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="de52d-126">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-126">Method</span></span> |
| [<span data-ttu-id="de52d-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="de52d-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="de52d-128">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-128">Method</span></span> |
| [<span data-ttu-id="de52d-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="de52d-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="de52d-130">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-130">Method</span></span> |
| [<span data-ttu-id="de52d-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="de52d-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="de52d-132">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-132">Method</span></span> |
| [<span data-ttu-id="de52d-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="de52d-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="de52d-134">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-134">Method</span></span> |
| [<span data-ttu-id="de52d-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="de52d-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="de52d-136">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-136">Method</span></span> |
| [<span data-ttu-id="de52d-137">Office.context.mailbox.displaynewmessageform</span><span class="sxs-lookup"><span data-stu-id="de52d-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="de52d-138">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-138">Method</span></span> |
| [<span data-ttu-id="de52d-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="de52d-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="de52d-140">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-140">Method</span></span> |
| [<span data-ttu-id="de52d-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="de52d-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="de52d-142">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-142">Method</span></span> |
| [<span data-ttu-id="de52d-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="de52d-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="de52d-144">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-144">Method</span></span> |
| [<span data-ttu-id="de52d-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="de52d-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="de52d-146">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-146">Method</span></span> |
| [<span data-ttu-id="de52d-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="de52d-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="de52d-148">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="de52d-149">命名空间</span><span class="sxs-lookup"><span data-stu-id="de52d-149">Namespaces</span></span>

<span data-ttu-id="de52d-150">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="de52d-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="de52d-151">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="de52d-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="de52d-152">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="de52d-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="de52d-153">Members</span><span class="sxs-lookup"><span data-stu-id="de52d-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="de52d-154">ewsUrl：String</span><span class="sxs-lookup"><span data-stu-id="de52d-154">ewsUrl: String</span></span>

<span data-ttu-id="de52d-p101">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="de52d-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-157">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="de52d-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="de52d-p102">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="de52d-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="de52d-160">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="de52d-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="de52d-p103">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="de52d-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="de52d-163">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-163">Type</span></span>

*   <span data-ttu-id="de52d-164">String</span><span class="sxs-lookup"><span data-stu-id="de52d-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="de52d-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-165">Requirements</span></span>

|<span data-ttu-id="de52d-166">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-166">Requirement</span></span>| <span data-ttu-id="de52d-167">值</span><span class="sxs-lookup"><span data-stu-id="de52d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-168">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-169">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-169">1.0</span></span>|
|[<span data-ttu-id="de52d-170">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-171">ReadItem</span></span>|
|[<span data-ttu-id="de52d-172">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-173">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="de52d-174">restUrl：String</span><span class="sxs-lookup"><span data-stu-id="de52d-174">restUrl: String</span></span>

<span data-ttu-id="de52d-175">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="de52d-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="de52d-176">`restUrl` 值可用于对用户邮箱进行 [REST API](/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="de52d-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="de52d-177">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-177">Type</span></span>

*   <span data-ttu-id="de52d-178">String</span><span class="sxs-lookup"><span data-stu-id="de52d-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="de52d-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-179">Requirements</span></span>

|<span data-ttu-id="de52d-180">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-180">Requirement</span></span>| <span data-ttu-id="de52d-181">值</span><span class="sxs-lookup"><span data-stu-id="de52d-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-183">1.5</span><span class="sxs-lookup"><span data-stu-id="de52d-183">1.5</span></span> |
|[<span data-ttu-id="de52d-184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-185">ReadItem</span></span>|
|[<span data-ttu-id="de52d-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-187">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="de52d-188">方法</span><span class="sxs-lookup"><span data-stu-id="de52d-188">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="de52d-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="de52d-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="de52d-190">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="de52d-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="de52d-191">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="de52d-191">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-192">参数</span><span class="sxs-lookup"><span data-stu-id="de52d-192">Parameters</span></span>

| <span data-ttu-id="de52d-193">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-193">Name</span></span> | <span data-ttu-id="de52d-194">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-194">Type</span></span> | <span data-ttu-id="de52d-195">属性</span><span class="sxs-lookup"><span data-stu-id="de52d-195">Attributes</span></span> | <span data-ttu-id="de52d-196">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="de52d-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="de52d-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="de52d-198">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="de52d-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="de52d-199">函数</span><span class="sxs-lookup"><span data-stu-id="de52d-199">Function</span></span> || <span data-ttu-id="de52d-p104">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="de52d-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="de52d-203">Object</span><span class="sxs-lookup"><span data-stu-id="de52d-203">Object</span></span> | <span data-ttu-id="de52d-204">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-204">&lt;optional&gt;</span></span> | <span data-ttu-id="de52d-205">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="de52d-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="de52d-206">对象</span><span class="sxs-lookup"><span data-stu-id="de52d-206">Object</span></span> | <span data-ttu-id="de52d-207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-207">&lt;optional&gt;</span></span> | <span data-ttu-id="de52d-208">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="de52d-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="de52d-209">函数</span><span class="sxs-lookup"><span data-stu-id="de52d-209">function</span></span>| <span data-ttu-id="de52d-210">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-210">&lt;optional&gt;</span></span>|<span data-ttu-id="de52d-211">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="de52d-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-212">Requirements</span></span>

|<span data-ttu-id="de52d-213">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-213">Requirement</span></span>| <span data-ttu-id="de52d-214">值</span><span class="sxs-lookup"><span data-stu-id="de52d-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-215">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-216">1.5</span><span class="sxs-lookup"><span data-stu-id="de52d-216">1.5</span></span> |
|[<span data-ttu-id="de52d-217">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-218">ReadItem</span></span> |
|[<span data-ttu-id="de52d-219">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-220">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="de52d-221">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-221">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="de52d-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="de52d-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="de52d-223">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="de52d-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-224">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="de52d-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="de52d-p105">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="de52d-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-227">Parameters</span><span class="sxs-lookup"><span data-stu-id="de52d-227">Parameters</span></span>

|<span data-ttu-id="de52d-228">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-228">Name</span></span>| <span data-ttu-id="de52d-229">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-229">Type</span></span>| <span data-ttu-id="de52d-230">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="de52d-231">字符串</span><span class="sxs-lookup"><span data-stu-id="de52d-231">String</span></span>|<span data-ttu-id="de52d-232">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="de52d-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="de52d-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="de52d-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="de52d-234">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="de52d-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-235">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-235">Requirements</span></span>

|<span data-ttu-id="de52d-236">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-236">Requirement</span></span>| <span data-ttu-id="de52d-237">值</span><span class="sxs-lookup"><span data-stu-id="de52d-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-238">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-239">1.3</span><span class="sxs-lookup"><span data-stu-id="de52d-239">1.3</span></span>|
|[<span data-ttu-id="de52d-240">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-241">受限</span><span class="sxs-lookup"><span data-stu-id="de52d-241">Restricted</span></span>|
|[<span data-ttu-id="de52d-242">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-243">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="de52d-244">返回：</span><span class="sxs-lookup"><span data-stu-id="de52d-244">Returns:</span></span>

<span data-ttu-id="de52d-245">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="de52d-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="de52d-246">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a><span data-ttu-id="de52d-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="de52d-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="de52d-248">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="de52d-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="de52d-p106">Outlook 桌面版或 Outlook 网页版邮件应用可以对日期和时间使用不同的时区。Outlook 桌面版使用客户端计算机时区；Outlook 网页版使用 Exchange 管理中心 (EAC) 中设置的时区。你应处理日期和时间值，让用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="de52d-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="de52d-p107">如果邮件应用是在 Outlook 桌面版客户端中运行，`convertToLocalClientTime` 方法返回值设置为客户端计算机时区的字典对象。如果邮件应用是在 Outlook 网页版中运行，`convertToLocalClientTime` 方法返回值设置为 EAC 中指定时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="de52d-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-254">Parameters</span><span class="sxs-lookup"><span data-stu-id="de52d-254">Parameters</span></span>

|<span data-ttu-id="de52d-255">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-255">Name</span></span>| <span data-ttu-id="de52d-256">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-256">Type</span></span>| <span data-ttu-id="de52d-257">描述</span><span class="sxs-lookup"><span data-stu-id="de52d-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="de52d-258">日期</span><span class="sxs-lookup"><span data-stu-id="de52d-258">Date</span></span>|<span data-ttu-id="de52d-259">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="de52d-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-260">Requirements</span></span>

|<span data-ttu-id="de52d-261">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-261">Requirement</span></span>| <span data-ttu-id="de52d-262">值</span><span class="sxs-lookup"><span data-stu-id="de52d-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-264">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-264">1.0</span></span>|
|[<span data-ttu-id="de52d-265">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-266">ReadItem</span></span>|
|[<span data-ttu-id="de52d-267">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-268">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="de52d-269">返回：</span><span class="sxs-lookup"><span data-stu-id="de52d-269">Returns:</span></span>

<span data-ttu-id="de52d-270">类型：[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="de52d-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="de52d-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="de52d-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="de52d-272">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="de52d-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-273">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="de52d-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="de52d-p108">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="de52d-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-276">Parameters</span><span class="sxs-lookup"><span data-stu-id="de52d-276">Parameters</span></span>

|<span data-ttu-id="de52d-277">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-277">Name</span></span>| <span data-ttu-id="de52d-278">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-278">Type</span></span>| <span data-ttu-id="de52d-279">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="de52d-280">String</span><span class="sxs-lookup"><span data-stu-id="de52d-280">String</span></span>|<span data-ttu-id="de52d-281">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="de52d-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="de52d-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="de52d-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="de52d-283">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="de52d-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-284">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-284">Requirements</span></span>

|<span data-ttu-id="de52d-285">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-285">Requirement</span></span>| <span data-ttu-id="de52d-286">值</span><span class="sxs-lookup"><span data-stu-id="de52d-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-287">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-288">1.3</span><span class="sxs-lookup"><span data-stu-id="de52d-288">1.3</span></span>|
|[<span data-ttu-id="de52d-289">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-290">受限</span><span class="sxs-lookup"><span data-stu-id="de52d-290">Restricted</span></span>|
|[<span data-ttu-id="de52d-291">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-292">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="de52d-293">返回：</span><span class="sxs-lookup"><span data-stu-id="de52d-293">Returns:</span></span>

<span data-ttu-id="de52d-294">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="de52d-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="de52d-295">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="de52d-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="de52d-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="de52d-297">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="de52d-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="de52d-298">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="de52d-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-299">Parameters</span><span class="sxs-lookup"><span data-stu-id="de52d-299">Parameters</span></span>

|<span data-ttu-id="de52d-300">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-300">Name</span></span>| <span data-ttu-id="de52d-301">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-301">Type</span></span>| <span data-ttu-id="de52d-302">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="de52d-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="de52d-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|<span data-ttu-id="de52d-304">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="de52d-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-305">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-305">Requirements</span></span>

|<span data-ttu-id="de52d-306">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-306">Requirement</span></span>| <span data-ttu-id="de52d-307">值</span><span class="sxs-lookup"><span data-stu-id="de52d-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-309">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-309">1.0</span></span>|
|[<span data-ttu-id="de52d-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-311">ReadItem</span></span>|
|[<span data-ttu-id="de52d-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-313">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="de52d-314">返回：</span><span class="sxs-lookup"><span data-stu-id="de52d-314">Returns:</span></span>

<span data-ttu-id="de52d-315">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="de52d-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="de52d-316">键入：日期</span><span class="sxs-lookup"><span data-stu-id="de52d-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="de52d-317">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-317">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="de52d-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="de52d-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="de52d-319">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="de52d-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-320">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="de52d-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="de52d-321">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="de52d-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="de52d-p109">在 Mac 版 Outlook 中，可以使用此方法来显示不属于重复时序的一个约会，或显示重复时序的主约会，但无法显示重复时序的实例。这是因为在 Mac 版 Outlook 中，无法访问重复时序的实例属性（包括项 ID）。</span><span class="sxs-lookup"><span data-stu-id="de52d-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="de52d-324">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="de52d-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="de52d-325">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="de52d-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-326">Parameters</span><span class="sxs-lookup"><span data-stu-id="de52d-326">Parameters</span></span>

|<span data-ttu-id="de52d-327">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-327">Name</span></span>| <span data-ttu-id="de52d-328">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-328">Type</span></span>| <span data-ttu-id="de52d-329">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="de52d-330">String</span><span class="sxs-lookup"><span data-stu-id="de52d-330">String</span></span>|<span data-ttu-id="de52d-331">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="de52d-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-332">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-332">Requirements</span></span>

|<span data-ttu-id="de52d-333">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-333">Requirement</span></span>| <span data-ttu-id="de52d-334">值</span><span class="sxs-lookup"><span data-stu-id="de52d-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-335">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-336">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-336">1.0</span></span>|
|[<span data-ttu-id="de52d-337">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-338">ReadItem</span></span>|
|[<span data-ttu-id="de52d-339">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-340">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="de52d-341">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="de52d-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="de52d-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="de52d-343">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="de52d-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-344">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="de52d-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="de52d-345">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="de52d-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="de52d-346">在 Outlook 网页版中，此方法仅在窗体正文的字符数小于或等于 32KB 时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="de52d-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="de52d-347">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="de52d-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="de52d-p110">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="de52d-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-350">参数</span><span class="sxs-lookup"><span data-stu-id="de52d-350">Parameters</span></span>

|<span data-ttu-id="de52d-351">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-351">Name</span></span>| <span data-ttu-id="de52d-352">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-352">Type</span></span>| <span data-ttu-id="de52d-353">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="de52d-354">字符串</span><span class="sxs-lookup"><span data-stu-id="de52d-354">String</span></span>|<span data-ttu-id="de52d-355">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="de52d-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-356">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-356">Requirements</span></span>

|<span data-ttu-id="de52d-357">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-357">Requirement</span></span>| <span data-ttu-id="de52d-358">值</span><span class="sxs-lookup"><span data-stu-id="de52d-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-359">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-360">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-360">1.0</span></span>|
|[<span data-ttu-id="de52d-361">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-362">ReadItem</span></span>|
|[<span data-ttu-id="de52d-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-364">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="de52d-365">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="de52d-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="de52d-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="de52d-367">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="de52d-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-368">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="de52d-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="de52d-p111">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="de52d-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="de52d-p112">在 Outlook 网页版和移动设备版中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，此方法显示包含“保存”\*\*\*\* 按钮的窗体。如果你已指定与会者，窗体包含与会者和“发送”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="de52d-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="de52d-p113">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="de52d-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="de52d-376">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="de52d-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-377">参数</span><span class="sxs-lookup"><span data-stu-id="de52d-377">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-378">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="de52d-378">All parameters are optional.</span></span>

|<span data-ttu-id="de52d-379">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-379">Name</span></span>| <span data-ttu-id="de52d-380">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-380">Type</span></span>| <span data-ttu-id="de52d-381">描述</span><span class="sxs-lookup"><span data-stu-id="de52d-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="de52d-382">对象</span><span class="sxs-lookup"><span data-stu-id="de52d-382">Object</span></span> | <span data-ttu-id="de52d-383">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="de52d-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="de52d-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="de52d-p114">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="de52d-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="de52d-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="de52d-p115">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="de52d-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="de52d-390">日期</span><span class="sxs-lookup"><span data-stu-id="de52d-390">Date</span></span> | <span data-ttu-id="de52d-391">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="de52d-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="de52d-392">Date</span><span class="sxs-lookup"><span data-stu-id="de52d-392">Date</span></span> | <span data-ttu-id="de52d-393">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="de52d-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="de52d-394">字符串</span><span class="sxs-lookup"><span data-stu-id="de52d-394">String</span></span> | <span data-ttu-id="de52d-p116">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="de52d-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="de52d-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="de52d-p117">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="de52d-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="de52d-400">String</span><span class="sxs-lookup"><span data-stu-id="de52d-400">String</span></span> | <span data-ttu-id="de52d-p118">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="de52d-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="de52d-403">字符串</span><span class="sxs-lookup"><span data-stu-id="de52d-403">String</span></span> | <span data-ttu-id="de52d-p119">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="de52d-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="de52d-406">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-406">Requirements</span></span>

|<span data-ttu-id="de52d-407">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-407">Requirement</span></span>| <span data-ttu-id="de52d-408">值</span><span class="sxs-lookup"><span data-stu-id="de52d-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-409">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-410">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-410">1.0</span></span>|
|[<span data-ttu-id="de52d-411">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-412">ReadItem</span></span>|
|[<span data-ttu-id="de52d-413">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-414">阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="de52d-415">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-415">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="de52d-416">Office.context.mailbox.displaynewmessageform （参数）</span><span class="sxs-lookup"><span data-stu-id="de52d-416">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="de52d-417">显示用于创建新邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="de52d-417">Displays a form for creating a new message.</span></span>

<span data-ttu-id="de52d-418">`displayNewMessageForm`方法将打开一个窗体，使用户可以创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="de52d-418">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="de52d-419">如果指定了参数，则将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="de52d-419">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="de52d-420">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="de52d-420">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-421">参数</span><span class="sxs-lookup"><span data-stu-id="de52d-421">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-422">所有参数都是可选的。</span><span class="sxs-lookup"><span data-stu-id="de52d-422">All parameters are optional.</span></span>

|<span data-ttu-id="de52d-423">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-423">Name</span></span>| <span data-ttu-id="de52d-424">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-424">Type</span></span>| <span data-ttu-id="de52d-425">描述</span><span class="sxs-lookup"><span data-stu-id="de52d-425">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="de52d-426">对象</span><span class="sxs-lookup"><span data-stu-id="de52d-426">Object</span></span> | <span data-ttu-id="de52d-427">描述新邮件的参数的字典。</span><span class="sxs-lookup"><span data-stu-id="de52d-427">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="de52d-428">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-428">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="de52d-429">包含电子邮件地址的字符串数组，或包含 "收件人" `EmailAddressDetails`行中每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="de52d-429">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="de52d-430">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="de52d-430">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="de52d-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="de52d-432">包含电子邮件地址的字符串数组，或包含 "抄送" `EmailAddressDetails`行上每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="de52d-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="de52d-433">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="de52d-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="de52d-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="de52d-435">包含电子邮件地址的字符串数组，或包含 Bcc 行上`EmailAddressDetails`每个收件人的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="de52d-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="de52d-436">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="de52d-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="de52d-437">String</span><span class="sxs-lookup"><span data-stu-id="de52d-437">String</span></span> | <span data-ttu-id="de52d-438">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="de52d-438">A string containing the subject of the message.</span></span> <span data-ttu-id="de52d-439">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="de52d-439">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="de52d-440">String</span><span class="sxs-lookup"><span data-stu-id="de52d-440">String</span></span> | <span data-ttu-id="de52d-441">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="de52d-441">The HTML body of the message.</span></span> <span data-ttu-id="de52d-442">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="de52d-442">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="de52d-443">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-443">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="de52d-444">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="de52d-444">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="de52d-445">String</span><span class="sxs-lookup"><span data-stu-id="de52d-445">String</span></span> | <span data-ttu-id="de52d-p126">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="de52d-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="de52d-448">String</span><span class="sxs-lookup"><span data-stu-id="de52d-448">String</span></span> | <span data-ttu-id="de52d-449">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="de52d-449">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="de52d-450">String</span><span class="sxs-lookup"><span data-stu-id="de52d-450">String</span></span> | <span data-ttu-id="de52d-p127">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="de52d-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="de52d-453">布尔</span><span class="sxs-lookup"><span data-stu-id="de52d-453">Boolean</span></span> | <span data-ttu-id="de52d-p128">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="de52d-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="de52d-456">字符串</span><span class="sxs-lookup"><span data-stu-id="de52d-456">String</span></span> | <span data-ttu-id="de52d-457">仅在 `type` 设置为 `item` 时使用。</span><span class="sxs-lookup"><span data-stu-id="de52d-457">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="de52d-458">要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="de52d-458">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="de52d-459">字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="de52d-459">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="de52d-460">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-460">Requirements</span></span>

|<span data-ttu-id="de52d-461">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-461">Requirement</span></span>| <span data-ttu-id="de52d-462">值</span><span class="sxs-lookup"><span data-stu-id="de52d-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-463">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-464">1.6</span><span class="sxs-lookup"><span data-stu-id="de52d-464">1.6</span></span> |
|[<span data-ttu-id="de52d-465">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-465">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-466">ReadItem</span></span>|
|[<span data-ttu-id="de52d-467">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-467">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-468">阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-468">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="de52d-469">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-469">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="de52d-470">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="de52d-470">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="de52d-471">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="de52d-471">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="de52d-p130">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="de52d-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-474">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="de52d-474">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="de52d-475">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="de52d-475">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="de52d-476">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="de52d-476">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="de52d-477">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="de52d-477">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="de52d-478">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="de52d-478">**REST Tokens**</span></span>

<span data-ttu-id="de52d-p132">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="de52d-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="de52d-482">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="de52d-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="de52d-483">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="de52d-483">**EWS Tokens**</span></span>

<span data-ttu-id="de52d-p133">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="de52d-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="de52d-486">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="de52d-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="de52d-487">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="de52d-487">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="de52d-488">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="de52d-488">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="de52d-489">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="de52d-489">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-490">Parameters</span><span class="sxs-lookup"><span data-stu-id="de52d-490">Parameters</span></span>

|<span data-ttu-id="de52d-491">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-491">Name</span></span>| <span data-ttu-id="de52d-492">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-492">Type</span></span>| <span data-ttu-id="de52d-493">属性</span><span class="sxs-lookup"><span data-stu-id="de52d-493">Attributes</span></span>| <span data-ttu-id="de52d-494">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-494">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="de52d-495">Object</span><span class="sxs-lookup"><span data-stu-id="de52d-495">Object</span></span> | <span data-ttu-id="de52d-496">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-496">&lt;optional&gt;</span></span> | <span data-ttu-id="de52d-497">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="de52d-497">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="de52d-498">布尔值</span><span class="sxs-lookup"><span data-stu-id="de52d-498">Boolean</span></span> |  <span data-ttu-id="de52d-499">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-499">&lt;optional&gt;</span></span> | <span data-ttu-id="de52d-p135">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="de52d-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="de52d-502">Object</span><span class="sxs-lookup"><span data-stu-id="de52d-502">Object</span></span> |  <span data-ttu-id="de52d-503">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-503">&lt;optional&gt;</span></span> | <span data-ttu-id="de52d-504">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="de52d-504">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="de52d-505">函数</span><span class="sxs-lookup"><span data-stu-id="de52d-505">function</span></span>||<span data-ttu-id="de52d-506">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="de52d-506">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="de52d-507">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="de52d-507">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="de52d-508">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="de52d-508">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="de52d-509">错误</span><span class="sxs-lookup"><span data-stu-id="de52d-509">Errors</span></span>

|<span data-ttu-id="de52d-510">错误代码</span><span class="sxs-lookup"><span data-stu-id="de52d-510">Error code</span></span>|<span data-ttu-id="de52d-511">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-511">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="de52d-512">请求失败。</span><span class="sxs-lookup"><span data-stu-id="de52d-512">The request has failed.</span></span> <span data-ttu-id="de52d-513">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="de52d-513">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="de52d-514">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="de52d-514">The Exchange server returned an error.</span></span> <span data-ttu-id="de52d-515">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="de52d-515">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="de52d-516">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="de52d-516">The user is no longer connected to the network.</span></span> <span data-ttu-id="de52d-517">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="de52d-517">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-518">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-518">Requirements</span></span>

|<span data-ttu-id="de52d-519">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-519">Requirement</span></span>| <span data-ttu-id="de52d-520">值</span><span class="sxs-lookup"><span data-stu-id="de52d-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-521">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-522">1.5</span><span class="sxs-lookup"><span data-stu-id="de52d-522">1.5</span></span> |
|[<span data-ttu-id="de52d-523">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-523">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-524">ReadItem</span></span>|
|[<span data-ttu-id="de52d-525">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-525">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-526">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-526">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="de52d-527">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-527">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="de52d-528">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="de52d-528">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="de52d-529">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="de52d-529">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="de52d-p139">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="de52d-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="de52d-532">可以将令牌和附件标识符或项标识符传递到第三方系统。</span><span class="sxs-lookup"><span data-stu-id="de52d-532">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="de52d-533">第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作或 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，以返回附件或项目。</span><span class="sxs-lookup"><span data-stu-id="de52d-533">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="de52d-534">例如，可以创建远程服务来 [获取选定项目中的附件](/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="de52d-534">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="de52d-535">在阅读模式下调用 `getCallbackTokenAsync` 方法要求最低权限级别的 **ReadItem**。</span><span class="sxs-lookup"><span data-stu-id="de52d-535">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="de52d-536">在撰写模式调用 `getCallbackTokenAsync` 要求已保存该项目。</span><span class="sxs-lookup"><span data-stu-id="de52d-536">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="de52d-537">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法要求最低权限级别的 **ReadWriteItem**。</span><span class="sxs-lookup"><span data-stu-id="de52d-537">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-538">Parameters</span><span class="sxs-lookup"><span data-stu-id="de52d-538">Parameters</span></span>

|<span data-ttu-id="de52d-539">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-539">Name</span></span>| <span data-ttu-id="de52d-540">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-540">Type</span></span>| <span data-ttu-id="de52d-541">属性</span><span class="sxs-lookup"><span data-stu-id="de52d-541">Attributes</span></span>| <span data-ttu-id="de52d-542">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-542">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="de52d-543">function</span><span class="sxs-lookup"><span data-stu-id="de52d-543">function</span></span>||<span data-ttu-id="de52d-544">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="de52d-544">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="de52d-545">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="de52d-545">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="de52d-546">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="de52d-546">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="de52d-547">对象</span><span class="sxs-lookup"><span data-stu-id="de52d-547">Object</span></span>| <span data-ttu-id="de52d-548">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-548">&lt;optional&gt;</span></span>|<span data-ttu-id="de52d-549">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="de52d-549">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="de52d-550">错误</span><span class="sxs-lookup"><span data-stu-id="de52d-550">Errors</span></span>

|<span data-ttu-id="de52d-551">错误代码</span><span class="sxs-lookup"><span data-stu-id="de52d-551">Error code</span></span>|<span data-ttu-id="de52d-552">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-552">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="de52d-553">请求失败。</span><span class="sxs-lookup"><span data-stu-id="de52d-553">The request has failed.</span></span> <span data-ttu-id="de52d-554">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="de52d-554">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="de52d-555">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="de52d-555">The Exchange server returned an error.</span></span> <span data-ttu-id="de52d-556">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="de52d-556">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="de52d-557">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="de52d-557">The user is no longer connected to the network.</span></span> <span data-ttu-id="de52d-558">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="de52d-558">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-559">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-559">Requirements</span></span>

|<span data-ttu-id="de52d-560">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-560">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="de52d-561">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-562">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-562">1.0</span></span> | <span data-ttu-id="de52d-563">1.3</span><span class="sxs-lookup"><span data-stu-id="de52d-563">1.3</span></span> |
|[<span data-ttu-id="de52d-564">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-565">ReadItem</span></span> | <span data-ttu-id="de52d-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-566">ReadItem</span></span> |
|[<span data-ttu-id="de52d-567">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-568">阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-568">Read</span></span> | <span data-ttu-id="de52d-569">撰写</span><span class="sxs-lookup"><span data-stu-id="de52d-569">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="de52d-570">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-570">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="de52d-571">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="de52d-571">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="de52d-572">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="de52d-572">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="de52d-573">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="de52d-573">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-574">参数</span><span class="sxs-lookup"><span data-stu-id="de52d-574">Parameters</span></span>

|<span data-ttu-id="de52d-575">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-575">Name</span></span>| <span data-ttu-id="de52d-576">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-576">Type</span></span>| <span data-ttu-id="de52d-577">属性</span><span class="sxs-lookup"><span data-stu-id="de52d-577">Attributes</span></span>| <span data-ttu-id="de52d-578">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-578">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="de52d-579">function</span><span class="sxs-lookup"><span data-stu-id="de52d-579">function</span></span>||<span data-ttu-id="de52d-580">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="de52d-580">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="de52d-581">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="de52d-581">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="de52d-582">如果出现错误，则 `asyncResult.error` 和 `asyncResult.diagnostics` 属性可能会提供其他信息。</span><span class="sxs-lookup"><span data-stu-id="de52d-582">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="de52d-583">对象</span><span class="sxs-lookup"><span data-stu-id="de52d-583">Object</span></span>| <span data-ttu-id="de52d-584">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-584">&lt;optional&gt;</span></span>|<span data-ttu-id="de52d-585">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="de52d-585">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="de52d-586">错误</span><span class="sxs-lookup"><span data-stu-id="de52d-586">Errors</span></span>

|<span data-ttu-id="de52d-587">错误代码</span><span class="sxs-lookup"><span data-stu-id="de52d-587">Error code</span></span>|<span data-ttu-id="de52d-588">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-588">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="de52d-589">请求失败。</span><span class="sxs-lookup"><span data-stu-id="de52d-589">The request has failed.</span></span> <span data-ttu-id="de52d-590">请查看诊断对象，了解 HTTP 错误代码。</span><span class="sxs-lookup"><span data-stu-id="de52d-590">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="de52d-591">Exchange 服务器返回了错误。</span><span class="sxs-lookup"><span data-stu-id="de52d-591">The Exchange server returned an error.</span></span> <span data-ttu-id="de52d-592">请查看诊断对象，了解详细信息。</span><span class="sxs-lookup"><span data-stu-id="de52d-592">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="de52d-593">用户不再连接到网络。</span><span class="sxs-lookup"><span data-stu-id="de52d-593">The user is no longer connected to the network.</span></span> <span data-ttu-id="de52d-594">请检查网络连接并重试。</span><span class="sxs-lookup"><span data-stu-id="de52d-594">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-595">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-595">Requirements</span></span>

|<span data-ttu-id="de52d-596">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-596">Requirement</span></span>| <span data-ttu-id="de52d-597">值</span><span class="sxs-lookup"><span data-stu-id="de52d-597">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-598">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-598">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-599">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-599">1.0</span></span>|
|[<span data-ttu-id="de52d-600">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-600">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-601">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-601">ReadItem</span></span>|
|[<span data-ttu-id="de52d-602">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-602">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-603">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-603">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="de52d-604">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-604">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="de52d-605">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="de52d-605">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="de52d-606">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="de52d-606">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-607">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="de52d-607">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="de52d-608">在 iOS 版 Outlook 或 Android 版 Outlook 中</span><span class="sxs-lookup"><span data-stu-id="de52d-608">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="de52d-609">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="de52d-609">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="de52d-610">在这些情况下，加载项应该[使用 REST API](/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="de52d-610">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="de52d-611">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="de52d-611">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="de52d-612">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="de52d-612">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="de52d-613">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="de52d-613">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="de52d-614">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="de52d-614">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="de52d-p149">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="de52d-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="de52d-617">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="de52d-617">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="de52d-618">版本差异</span><span class="sxs-lookup"><span data-stu-id="de52d-618">Version differences</span></span>

<span data-ttu-id="de52d-619">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="de52d-619">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="de52d-p150">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="de52d-p150">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-623">Parameters</span><span class="sxs-lookup"><span data-stu-id="de52d-623">Parameters</span></span>

|<span data-ttu-id="de52d-624">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-624">Name</span></span>| <span data-ttu-id="de52d-625">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-625">Type</span></span>| <span data-ttu-id="de52d-626">属性</span><span class="sxs-lookup"><span data-stu-id="de52d-626">Attributes</span></span>| <span data-ttu-id="de52d-627">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-627">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="de52d-628">字符串</span><span class="sxs-lookup"><span data-stu-id="de52d-628">String</span></span>||<span data-ttu-id="de52d-629">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="de52d-629">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="de52d-630">函数</span><span class="sxs-lookup"><span data-stu-id="de52d-630">function</span></span>||<span data-ttu-id="de52d-631">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="de52d-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="de52d-632">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="de52d-632">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="de52d-633">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="de52d-633">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="de52d-634">对象</span><span class="sxs-lookup"><span data-stu-id="de52d-634">Object</span></span>| <span data-ttu-id="de52d-635">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-635">&lt;optional&gt;</span></span>|<span data-ttu-id="de52d-636">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="de52d-636">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-637">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-637">Requirements</span></span>

|<span data-ttu-id="de52d-638">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-638">Requirement</span></span>| <span data-ttu-id="de52d-639">值</span><span class="sxs-lookup"><span data-stu-id="de52d-639">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-640">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-640">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-641">1.0</span><span class="sxs-lookup"><span data-stu-id="de52d-641">1.0</span></span>|
|[<span data-ttu-id="de52d-642">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-642">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-643">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="de52d-643">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="de52d-644">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-644">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-645">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-645">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="de52d-646">示例</span><span class="sxs-lookup"><span data-stu-id="de52d-646">Example</span></span>

<span data-ttu-id="de52d-647">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="de52d-647">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="de52d-648">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="de52d-648">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="de52d-649">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="de52d-649">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="de52d-650">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="de52d-650">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="de52d-651">参数</span><span class="sxs-lookup"><span data-stu-id="de52d-651">Parameters</span></span>

| <span data-ttu-id="de52d-652">名称</span><span class="sxs-lookup"><span data-stu-id="de52d-652">Name</span></span> | <span data-ttu-id="de52d-653">类型</span><span class="sxs-lookup"><span data-stu-id="de52d-653">Type</span></span> | <span data-ttu-id="de52d-654">属性</span><span class="sxs-lookup"><span data-stu-id="de52d-654">Attributes</span></span> | <span data-ttu-id="de52d-655">说明</span><span class="sxs-lookup"><span data-stu-id="de52d-655">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="de52d-656">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="de52d-656">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="de52d-657">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="de52d-657">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="de52d-658">对象</span><span class="sxs-lookup"><span data-stu-id="de52d-658">Object</span></span> | <span data-ttu-id="de52d-659">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-659">&lt;optional&gt;</span></span> | <span data-ttu-id="de52d-660">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="de52d-660">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="de52d-661">对象</span><span class="sxs-lookup"><span data-stu-id="de52d-661">Object</span></span> | <span data-ttu-id="de52d-662">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-662">&lt;optional&gt;</span></span> | <span data-ttu-id="de52d-663">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="de52d-663">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="de52d-664">函数</span><span class="sxs-lookup"><span data-stu-id="de52d-664">function</span></span>| <span data-ttu-id="de52d-665">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="de52d-665">&lt;optional&gt;</span></span>|<span data-ttu-id="de52d-666">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="de52d-666">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="de52d-667">Requirements</span><span class="sxs-lookup"><span data-stu-id="de52d-667">Requirements</span></span>

|<span data-ttu-id="de52d-668">要求</span><span class="sxs-lookup"><span data-stu-id="de52d-668">Requirement</span></span>| <span data-ttu-id="de52d-669">值</span><span class="sxs-lookup"><span data-stu-id="de52d-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="de52d-670">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="de52d-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de52d-671">1.5</span><span class="sxs-lookup"><span data-stu-id="de52d-671">1.5</span></span> |
|[<span data-ttu-id="de52d-672">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="de52d-672">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de52d-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de52d-673">ReadItem</span></span> |
|[<span data-ttu-id="de52d-674">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="de52d-674">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de52d-675">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="de52d-675">Compose or Read</span></span>|
