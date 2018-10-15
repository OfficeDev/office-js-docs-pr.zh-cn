
# <a name="mailbox"></a><span data-ttu-id="23ad3-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="23ad3-101">mailbox</span></span>

### <span data-ttu-id="23ad3-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="23ad3-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="23ad3-104">为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="23ad3-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="23ad3-105">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-105">Requirements</span></span>

|<span data-ttu-id="23ad3-106">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-106">Requirement</span></span>| <span data-ttu-id="23ad3-107">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-108">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-109">1.0</span></span>|
|[<span data-ttu-id="23ad3-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="23ad3-111">Restricted</span></span>|
|[<span data-ttu-id="23ad3-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="23ad3-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-114">Members and methods</span></span>

| <span data-ttu-id="23ad3-115">成员</span><span class="sxs-lookup"><span data-stu-id="23ad3-115">Member</span></span> | <span data-ttu-id="23ad3-116">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="23ad3-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="23ad3-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="23ad3-118">成员</span><span class="sxs-lookup"><span data-stu-id="23ad3-118">Member</span></span> |
| [<span data-ttu-id="23ad3-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="23ad3-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="23ad3-120">成员</span><span class="sxs-lookup"><span data-stu-id="23ad3-120">Member</span></span> |
| [<span data-ttu-id="23ad3-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="23ad3-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="23ad3-122">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-122">Method</span></span> |
| [<span data-ttu-id="23ad3-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="23ad3-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="23ad3-124">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-124">Method</span></span> |
| [<span data-ttu-id="23ad3-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="23ad3-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) | <span data-ttu-id="23ad3-126">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-126">Method</span></span> |
| [<span data-ttu-id="23ad3-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="23ad3-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="23ad3-128">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-128">Method</span></span> |
| [<span data-ttu-id="23ad3-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="23ad3-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="23ad3-130">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-130">Method</span></span> |
| [<span data-ttu-id="23ad3-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="23ad3-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="23ad3-132">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-132">Method</span></span> |
| [<span data-ttu-id="23ad3-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="23ad3-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="23ad3-134">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-134">Method</span></span> |
| [<span data-ttu-id="23ad3-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="23ad3-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="23ad3-136">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-136">Method</span></span> |
| [<span data-ttu-id="23ad3-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="23ad3-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="23ad3-138">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-138">Method</span></span> |
| [<span data-ttu-id="23ad3-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="23ad3-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="23ad3-140">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-140">Method</span></span> |
| [<span data-ttu-id="23ad3-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="23ad3-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="23ad3-142">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-142">Method</span></span> |
| [<span data-ttu-id="23ad3-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="23ad3-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="23ad3-144">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-144">Method</span></span> |
| [<span data-ttu-id="23ad3-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="23ad3-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="23ad3-146">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="23ad3-147">命名空间</span><span class="sxs-lookup"><span data-stu-id="23ad3-147">Namespaces</span></span>

<span data-ttu-id="23ad3-148">[诊断](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="23ad3-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="23ad3-149">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 加载项中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="23ad3-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="23ad3-150">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户信息。</span><span class="sxs-lookup"><span data-stu-id="23ad3-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="23ad3-151">成员</span><span class="sxs-lookup"><span data-stu-id="23ad3-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="23ad3-152">ewsUrl： 字符串</span><span class="sxs-lookup"><span data-stu-id="23ad3-152">ewsUrl :String</span></span>

<span data-ttu-id="23ad3-p102">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 端点 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-155">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="23ad3-155">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="23ad3-p103">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="23ad3-158">应用必须在其清单中指定拥有 **ReadItem** 权限，才能在阅读模式中调用 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="23ad3-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="23ad3-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="23ad3-161">类型：</span><span class="sxs-lookup"><span data-stu-id="23ad3-161">Type:</span></span>

*   <span data-ttu-id="23ad3-162">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="23ad3-163">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-163">Requirements</span></span>

|<span data-ttu-id="23ad3-164">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-164">Requirement</span></span>| <span data-ttu-id="23ad3-165">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-166">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-167">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-167">1.0</span></span>|
|[<span data-ttu-id="23ad3-168">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-169">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="23ad3-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="23ad3-172">restUrl :String</span></span>

<span data-ttu-id="23ad3-173">获取此电子邮件帐户的 REST 端点 URL。</span><span class="sxs-lookup"><span data-stu-id="23ad3-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="23ad3-174">`restUrl` 值可用于对用户邮箱进行 [REST API](https://docs.microsoft.com/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="23ad3-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="23ad3-175">应用必须在其清单中指定拥有 **ReadItem** 权限，才能在阅读模式中调用 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="23ad3-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="23ad3-p105">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="23ad3-178">类型：</span><span class="sxs-lookup"><span data-stu-id="23ad3-178">Type:</span></span>

*   <span data-ttu-id="23ad3-179">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="23ad3-180">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-180">Requirements</span></span>

|<span data-ttu-id="23ad3-181">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-181">Requirement</span></span>| <span data-ttu-id="23ad3-182">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-183">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-183">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-184">1.5</span><span class="sxs-lookup"><span data-stu-id="23ad3-184">1.5</span></span> |
|[<span data-ttu-id="23ad3-185">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-186">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-187">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-188">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="23ad3-189">方法</span><span class="sxs-lookup"><span data-stu-id="23ad3-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="23ad3-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="23ad3-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="23ad3-191">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="23ad3-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="23ad3-192">目前，支持的事件类型是 `Office.EventType.ItemChanged` 和 `Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="23ad3-192">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-193">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-193">Parameters:</span></span>

| <span data-ttu-id="23ad3-194">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-194">Name</span></span> | <span data-ttu-id="23ad3-195">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-195">Type</span></span> | <span data-ttu-id="23ad3-196">属性</span><span class="sxs-lookup"><span data-stu-id="23ad3-196">Attributes</span></span> | <span data-ttu-id="23ad3-197">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="23ad3-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="23ad3-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="23ad3-199">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="23ad3-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="23ad3-200">Function</span><span class="sxs-lookup"><span data-stu-id="23ad3-200">Function</span></span> || <span data-ttu-id="23ad3-p106">用于处理事件的函数。此函数必须接受单个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="23ad3-204">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-204">Object</span></span> | <span data-ttu-id="23ad3-205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-205">&lt;optional&gt;</span></span> | <span data-ttu-id="23ad3-206">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="23ad3-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="23ad3-207">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-207">Object</span></span> | <span data-ttu-id="23ad3-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-208">&lt;optional&gt;</span></span> | <span data-ttu-id="23ad3-209">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="23ad3-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="23ad3-210">函数</span><span class="sxs-lookup"><span data-stu-id="23ad3-210">function</span></span>| <span data-ttu-id="23ad3-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-211">&lt;optional&gt;</span></span>|<span data-ttu-id="23ad3-212">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="23ad3-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-213">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-213">Requirements</span></span>

|<span data-ttu-id="23ad3-214">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-214">Requirement</span></span>| <span data-ttu-id="23ad3-215">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-216">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-217">1.5</span><span class="sxs-lookup"><span data-stu-id="23ad3-217">1.5</span></span> |
|[<span data-ttu-id="23ad3-218">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-219">ReadItem</span></span> |
|[<span data-ttu-id="23ad3-220">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-221">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-222">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-222">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="23ad3-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="23ad3-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="23ad3-224">将适用 REST 格式化的项目 ID 转换为 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="23ad3-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-225">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="23ad3-225">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="23ad3-p107">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](http://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-228">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-228">Parameters:</span></span>

|<span data-ttu-id="23ad3-229">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-229">Name</span></span>| <span data-ttu-id="23ad3-230">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-230">Type</span></span>| <span data-ttu-id="23ad3-231">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="23ad3-232">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-232">String</span></span>|<span data-ttu-id="23ad3-233">适用 Outlook REST API 进行格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="23ad3-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="23ad3-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="23ad3-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="23ad3-235">值指示用于检索项目 ID 的 Outlook REST API 版本。</span><span class="sxs-lookup"><span data-stu-id="23ad3-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-236">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-236">Requirements</span></span>

|<span data-ttu-id="23ad3-237">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-237">Requirement</span></span>| <span data-ttu-id="23ad3-238">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-239">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-240">1.3</span><span class="sxs-lookup"><span data-stu-id="23ad3-240">1.3</span></span>|
|[<span data-ttu-id="23ad3-241">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-242">Restricted</span><span class="sxs-lookup"><span data-stu-id="23ad3-242">Restricted</span></span>|
|[<span data-ttu-id="23ad3-243">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-244">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="23ad3-245">返回：</span><span class="sxs-lookup"><span data-stu-id="23ad3-245">Returns:</span></span>

<span data-ttu-id="23ad3-246">类型：String</span><span class="sxs-lookup"><span data-stu-id="23ad3-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="23ad3-247">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="23ad3-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="23ad3-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="23ad3-249">获取包含以本地客户端时间表示时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="23ad3-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="23ad3-p108">Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="23ad3-p109">如果在 Outlook 中运行邮件应用程序，`convertToLocalClientTime` 方法将返回多个值设置为客户端计算机时区的字典对象。如果在 Outlook Web App 中运行邮件应用程序，`convertToLocalClientTime` 方法将返回多个值设置为 EAC 中指定的时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-255">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-255">Parameters:</span></span>

|<span data-ttu-id="23ad3-256">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-256">Name</span></span>| <span data-ttu-id="23ad3-257">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-257">Type</span></span>| <span data-ttu-id="23ad3-258">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="23ad3-259">Date</span><span class="sxs-lookup"><span data-stu-id="23ad3-259">Date</span></span>|<span data-ttu-id="23ad3-260">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="23ad3-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-261">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-261">Requirements</span></span>

|<span data-ttu-id="23ad3-262">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-262">Requirement</span></span>| <span data-ttu-id="23ad3-263">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-264">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-265">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-265">1.0</span></span>|
|[<span data-ttu-id="23ad3-266">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-267">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-268">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-269">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="23ad3-270">返回：</span><span class="sxs-lookup"><span data-stu-id="23ad3-270">Returns:</span></span>

<span data-ttu-id="23ad3-271">返回：LocalClientTime[ ](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="23ad3-271">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="23ad3-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="23ad3-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="23ad3-273">将适用 EWS 格式化的项目 ID 转换为 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="23ad3-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-274">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="23ad3-274">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="23ad3-p110">通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用与 REST API 不同的格式（例如 [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](http://graph.microsoft.io/)）。`convertToRestId` 方法将适用 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-277">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-277">Parameters:</span></span>

|<span data-ttu-id="23ad3-278">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-278">Name</span></span>| <span data-ttu-id="23ad3-279">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-279">Type</span></span>| <span data-ttu-id="23ad3-280">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="23ad3-281">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-281">String</span></span>|<span data-ttu-id="23ad3-282">适用于 Exchange Web 服务 (EWS) 进行格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="23ad3-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="23ad3-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="23ad3-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="23ad3-284">值指示转换的 ID 所使用的 Outlook REST API 版本。</span><span class="sxs-lookup"><span data-stu-id="23ad3-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-285">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-285">Requirements</span></span>

|<span data-ttu-id="23ad3-286">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-286">Requirement</span></span>| <span data-ttu-id="23ad3-287">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-288">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-289">1.3</span><span class="sxs-lookup"><span data-stu-id="23ad3-289">1.3</span></span>|
|[<span data-ttu-id="23ad3-290">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-291">Restricted</span><span class="sxs-lookup"><span data-stu-id="23ad3-291">Restricted</span></span>|
|[<span data-ttu-id="23ad3-292">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-293">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="23ad3-294">返回：</span><span class="sxs-lookup"><span data-stu-id="23ad3-294">Returns:</span></span>

<span data-ttu-id="23ad3-295">类型：String</span><span class="sxs-lookup"><span data-stu-id="23ad3-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="23ad3-296">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="23ad3-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="23ad3-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="23ad3-298">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="23ad3-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="23ad3-299">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="23ad3-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-300">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-300">Parameters:</span></span>

|<span data-ttu-id="23ad3-301">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-301">Name</span></span>| <span data-ttu-id="23ad3-302">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-302">Type</span></span>| <span data-ttu-id="23ad3-303">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="23ad3-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="23ad3-304">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="23ad3-305">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="23ad3-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-306">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-306">Requirements</span></span>

|<span data-ttu-id="23ad3-307">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-307">Requirement</span></span>| <span data-ttu-id="23ad3-308">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-309">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-310">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-310">1.0</span></span>|
|[<span data-ttu-id="23ad3-311">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-312">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-313">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-314">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="23ad3-315">返回：</span><span class="sxs-lookup"><span data-stu-id="23ad3-315">Returns:</span></span>

<span data-ttu-id="23ad3-316">以UTC格式表示时间的 Date 对象</span><span class="sxs-lookup"><span data-stu-id="23ad3-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="23ad3-317">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="23ad3-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="23ad3-318">Date</span><span class="sxs-lookup"><span data-stu-id="23ad3-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="23ad3-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="23ad3-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="23ad3-320">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="23ad3-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-321">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="23ad3-321">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="23ad3-322">`displayAppointmentForm` 方法将在桌面新窗口中或移动设备对话框中打开现有的日历约会。</span><span class="sxs-lookup"><span data-stu-id="23ad3-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="23ad3-p111">在 Outlook for Mac 中，您可以使用此方法来显示非重复性的单个约会，或显示重复系列中的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="23ad3-325">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="23ad3-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="23ad3-326">如果指定的项目标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="23ad3-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-327">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-327">Parameters:</span></span>

|<span data-ttu-id="23ad3-328">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-328">Name</span></span>| <span data-ttu-id="23ad3-329">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-329">Type</span></span>| <span data-ttu-id="23ad3-330">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="23ad3-331">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-331">String</span></span>|<span data-ttu-id="23ad3-332">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="23ad3-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-333">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-333">Requirements</span></span>

|<span data-ttu-id="23ad3-334">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-334">Requirement</span></span>| <span data-ttu-id="23ad3-335">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-336">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-337">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-337">1.0</span></span>|
|[<span data-ttu-id="23ad3-338">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-339">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-340">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-341">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-342">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="23ad3-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="23ad3-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="23ad3-344">显示一封现有邮件。</span><span class="sxs-lookup"><span data-stu-id="23ad3-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-345">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="23ad3-345">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="23ad3-346">在 Outlook Web App 中，`displayMessageForm` 方法将在桌面新窗口中或移动设备对话框中打开一封现有邮件。</span><span class="sxs-lookup"><span data-stu-id="23ad3-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="23ad3-347">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="23ad3-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="23ad3-348">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="23ad3-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="23ad3-p112">请勿使用 `displayMessageForm` 配合 `itemId` 表示约会。 使用 `displayAppointmentForm` 方法显示一个现有约会， 并 `displayNewAppointmentForm` 显示一个创建新约会的窗体。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-351">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-351">Parameters:</span></span>

|<span data-ttu-id="23ad3-352">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-352">Name</span></span>| <span data-ttu-id="23ad3-353">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-353">Type</span></span>| <span data-ttu-id="23ad3-354">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="23ad3-355">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-355">String</span></span>|<span data-ttu-id="23ad3-356">现有邮件的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="23ad3-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-357">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-357">Requirements</span></span>

|<span data-ttu-id="23ad3-358">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-358">Requirement</span></span>| <span data-ttu-id="23ad3-359">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-360">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-361">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-361">1.0</span></span>|
|[<span data-ttu-id="23ad3-362">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-363">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-364">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-365">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-366">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="23ad3-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="23ad3-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="23ad3-368">显示用于新建日历约会的窗体。</span><span class="sxs-lookup"><span data-stu-id="23ad3-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-369">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="23ad3-369">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="23ad3-p113">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充参数内容。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="23ad3-p114">在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="23ad3-p115">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="23ad3-377">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="23ad3-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-378">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-378">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-379">所有参数都是可选参数。</span><span class="sxs-lookup"><span data-stu-id="23ad3-379">Note: All parameters are optional.</span></span>

|<span data-ttu-id="23ad3-380">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-380">Name</span></span>| <span data-ttu-id="23ad3-381">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-381">Type</span></span>| <span data-ttu-id="23ad3-382">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="23ad3-383">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-383">Object</span></span> | <span data-ttu-id="23ad3-384">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="23ad3-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="23ad3-385">数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="23ad3-p116">包含电子邮件地址的字符串数组或包含约会的每个必需与会者 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="23ad3-388">数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="23ad3-p117">包含电子邮件地址的字符串数组或包含约会的每个可选与会者 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="23ad3-391">Date</span><span class="sxs-lookup"><span data-stu-id="23ad3-391">Date</span></span> | <span data-ttu-id="23ad3-392">指定约会开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="23ad3-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="23ad3-393">Date</span><span class="sxs-lookup"><span data-stu-id="23ad3-393">Date</span></span> | <span data-ttu-id="23ad3-394">指定约会的结束日期和时间的  对象。`Date`</span><span class="sxs-lookup"><span data-stu-id="23ad3-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="23ad3-395">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-395">String</span></span> | <span data-ttu-id="23ad3-p118">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="23ad3-398">数组.&lt;字符串&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="23ad3-p119">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="23ad3-401">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-401">String</span></span> | <span data-ttu-id="23ad3-p120">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="23ad3-404">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-404">String</span></span> | <span data-ttu-id="23ad3-p121">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="23ad3-407">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-407">Requirements</span></span>

|<span data-ttu-id="23ad3-408">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-408">Requirement</span></span>| <span data-ttu-id="23ad3-409">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-410">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-411">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-411">1.0</span></span>|
|[<span data-ttu-id="23ad3-412">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-413">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-414">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-415">阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-416">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-416">Example</span></span>

```
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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="23ad3-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="23ad3-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="23ad3-418">显示用于新建邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="23ad3-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="23ad3-419">`displayNewMessageForm` 方法将打开可让用户新建邮件的窗体。</span><span class="sxs-lookup"><span data-stu-id="23ad3-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="23ad3-420">如果指定了参数，将使用参数的内容自动填充邮件窗体字段。</span><span class="sxs-lookup"><span data-stu-id="23ad3-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="23ad3-421">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="23ad3-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-422">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-422">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-423">所有参数都是可选参数。</span><span class="sxs-lookup"><span data-stu-id="23ad3-423">Note: All parameters are optional.</span></span>

|<span data-ttu-id="23ad3-424">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-424">Name</span></span>| <span data-ttu-id="23ad3-425">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-425">Type</span></span>| <span data-ttu-id="23ad3-426">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="23ad3-427">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-427">Object</span></span> | <span data-ttu-id="23ad3-428">描述新邮件的参数字典。</span><span class="sxs-lookup"><span data-stu-id="23ad3-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="23ad3-429">数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="23ad3-430">包含电子邮件地址的字符串数组或包含收件人行上每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="23ad3-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="23ad3-431">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="23ad3-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="23ad3-432">数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="23ad3-433">包含电子邮件地址的字符串数组或包含抄送行上每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="23ad3-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="23ad3-434">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="23ad3-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="23ad3-435">数组.&lt;字符串&gt; | 数组.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="23ad3-436">包含电子邮件地址的字符串数组或包含密件抄送行上每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="23ad3-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="23ad3-437">数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="23ad3-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="23ad3-438">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-438">String</span></span> | <span data-ttu-id="23ad3-439">包含邮件主题的字符串。</span><span class="sxs-lookup"><span data-stu-id="23ad3-439">A string containing the subject of the message.</span></span> <span data-ttu-id="23ad3-440">字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="23ad3-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="23ad3-441">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-441">String</span></span> | <span data-ttu-id="23ad3-442">邮件的 HTML 正文。</span><span class="sxs-lookup"><span data-stu-id="23ad3-442">The HTML body of the message.</span></span> <span data-ttu-id="23ad3-443">正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="23ad3-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="23ad3-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="23ad3-445">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="23ad3-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="23ad3-446">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-446">String</span></span> | <span data-ttu-id="23ad3-p128">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="23ad3-449">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-449">String</span></span> | <span data-ttu-id="23ad3-450">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="23ad3-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="23ad3-451">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-451">String</span></span> | <span data-ttu-id="23ad3-p129">仅在 `type` 设置为 `file` 时才使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="23ad3-454">Boolean</span><span class="sxs-lookup"><span data-stu-id="23ad3-454">Boolean</span></span> | <span data-ttu-id="23ad3-p130">仅在 `type` 设置为 `file` 时才使用。如果为 `true`，表示将在邮件正文中嵌入显示附件，并且不应在附件列表中显示。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="23ad3-457">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-457">String</span></span> | <span data-ttu-id="23ad3-458">仅在 `type` 设置为 `item` 时才使用。</span><span class="sxs-lookup"><span data-stu-id="23ad3-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="23ad3-459">你想要附加到新邮件的现有电子邮件的 EWS 项目 id。</span><span class="sxs-lookup"><span data-stu-id="23ad3-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="23ad3-460">字符串最多达 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="23ad3-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="23ad3-461">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-461">Requirements</span></span>

|<span data-ttu-id="23ad3-462">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-462">Requirement</span></span>| <span data-ttu-id="23ad3-463">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-464">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-464">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-465">1.6</span><span class="sxs-lookup"><span data-stu-id="23ad3-465">-16</span></span> |
|[<span data-ttu-id="23ad3-466">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-466">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-467">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-468">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-468">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-469">阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-470">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-470">Example</span></span>

```
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="23ad3-471">getCallbackTokenAsync([选项] 回调)</span><span class="sxs-lookup"><span data-stu-id="23ad3-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="23ad3-472">获取一个包含用于调用 REST API 或 Exchange Web 服务令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="23ad3-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="23ad3-p132">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取不透明令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-475">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="23ad3-475">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="23ad3-476">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="23ad3-476">**REST Tokens**</span></span>

<span data-ttu-id="23ad3-p133">请求 REST 令牌 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非加载项在其清单中指定了 [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="23ad3-480">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="23ad3-480">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="23ad3-481">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="23ad3-481">**EWS Tokens**</span></span>

<span data-ttu-id="23ad3-p134">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="23ad3-484">加载项应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="23ad3-484">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-485">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-485">Parameters:</span></span>

|<span data-ttu-id="23ad3-486">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-486">Name</span></span>| <span data-ttu-id="23ad3-487">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-487">Type</span></span>| <span data-ttu-id="23ad3-488">属性</span><span class="sxs-lookup"><span data-stu-id="23ad3-488">Attributes</span></span>| <span data-ttu-id="23ad3-489">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-489">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="23ad3-490">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-490">Object</span></span> | <span data-ttu-id="23ad3-491">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-491">&lt;optional&gt;</span></span> | <span data-ttu-id="23ad3-492">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="23ad3-492">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="23ad3-493">Boolean</span><span class="sxs-lookup"><span data-stu-id="23ad3-493">Boolean</span></span> |  <span data-ttu-id="23ad3-494">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-494">&lt;optional&gt;</span></span> | <span data-ttu-id="23ad3-p135">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="23ad3-497">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-497">Object</span></span> |  <span data-ttu-id="23ad3-498">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-498">&lt;optional&gt;</span></span> | <span data-ttu-id="23ad3-499">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="23ad3-499">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="23ad3-500">function</span><span class="sxs-lookup"><span data-stu-id="23ad3-500">function</span></span>||<span data-ttu-id="23ad3-p136">方法完成后，通过单个参数调用 `callback` 参数中传递的函数， `asyncResult`, 是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象。令牌以 `asyncResult.value` 属性字符串形式提供。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-503">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-503">Requirements</span></span>

|<span data-ttu-id="23ad3-504">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-504">Requirement</span></span>| <span data-ttu-id="23ad3-505">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-506">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-507">1.5</span><span class="sxs-lookup"><span data-stu-id="23ad3-507">1.5</span></span> |
|[<span data-ttu-id="23ad3-508">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-509">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-510">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-511">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-511">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-512">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-512">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="23ad3-513">getCallbackTokenAsync(回调, [userContext])</span><span class="sxs-lookup"><span data-stu-id="23ad3-513">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="23ad3-514">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="23ad3-514">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="23ad3-p137">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取不透明令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="23ad3-p138">可以将令牌和附件标识符或项标识符传递至第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="23ad3-520">应用必须在其清单中指定拥有 **ReadItem** 权限，才能在阅读模式中调用 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="23ad3-520">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="23ad3-p139">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-523">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-523">Parameters:</span></span>

|<span data-ttu-id="23ad3-524">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-524">Name</span></span>| <span data-ttu-id="23ad3-525">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-525">Type</span></span>| <span data-ttu-id="23ad3-526">属性</span><span class="sxs-lookup"><span data-stu-id="23ad3-526">Attributes</span></span>| <span data-ttu-id="23ad3-527">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-527">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="23ad3-528">function</span><span class="sxs-lookup"><span data-stu-id="23ad3-528">function</span></span>||<span data-ttu-id="23ad3-p140">方法完成后，通过单个参数 `asyncResult`（这是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用 `callback` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="23ad3-531">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-531">Object</span></span>| <span data-ttu-id="23ad3-532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-532">&lt;optional&gt;</span></span>|<span data-ttu-id="23ad3-533">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="23ad3-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-534">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-534">Requirements</span></span>

|<span data-ttu-id="23ad3-535">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-535">Requirement</span></span>| <span data-ttu-id="23ad3-536">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-537">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-538">1.3</span><span class="sxs-lookup"><span data-stu-id="23ad3-538">1.3</span></span>|
|[<span data-ttu-id="23ad3-539">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-539">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-540">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-541">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-541">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-542">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-542">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-543">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-543">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="23ad3-544">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="23ad3-544">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="23ad3-545">获取用于标识用户和 Office 加载项的令牌。</span><span class="sxs-lookup"><span data-stu-id="23ad3-545">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="23ad3-546">`getUserIdentityTokenAsync` 方法返回可以用于标识和[在第三方系统上验证加载项和用户](https://docs.microsoft.com/outlook/add-ins/authentication)的令牌。</span><span class="sxs-lookup"><span data-stu-id="23ad3-546">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-547">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-547">Parameters:</span></span>

|<span data-ttu-id="23ad3-548">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-548">Name</span></span>| <span data-ttu-id="23ad3-549">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-549">Type</span></span>| <span data-ttu-id="23ad3-550">属性</span><span class="sxs-lookup"><span data-stu-id="23ad3-550">Attributes</span></span>| <span data-ttu-id="23ad3-551">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-551">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="23ad3-552">function</span><span class="sxs-lookup"><span data-stu-id="23ad3-552">function</span></span>||<span data-ttu-id="23ad3-553">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="23ad3-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="23ad3-554">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="23ad3-554">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="23ad3-555">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-555">Object</span></span>| <span data-ttu-id="23ad3-556">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-556">&lt;optional&gt;</span></span>|<span data-ttu-id="23ad3-557">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="23ad3-557">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-558">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-558">Requirements</span></span>

|<span data-ttu-id="23ad3-559">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-559">Requirement</span></span>| <span data-ttu-id="23ad3-560">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-561">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-562">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-562">1.0</span></span>|
|[<span data-ttu-id="23ad3-563">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="23ad3-564">ReadItem</span></span>|
|[<span data-ttu-id="23ad3-565">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-566">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-567">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-567">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="23ad3-568">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="23ad3-568">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="23ad3-569">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="23ad3-569">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-570">在以下方案中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="23ad3-570">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="23ad3-571">在 Outlook for iOS 或 Outlook for Android 中</span><span class="sxs-lookup"><span data-stu-id="23ad3-571">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="23ad3-572">在 Gmail 邮箱中加载加载项时</span><span class="sxs-lookup"><span data-stu-id="23ad3-572">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="23ad3-573">在这些情况下, 加载项应转而[使用 REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) 访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="23ad3-573">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="23ad3-574">`makeEwsRequestAsync` 方法代表加载项向 Exchange 发送 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="23ad3-574">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="23ad3-575">有关支持的 EWS 操作列表，请参阅 [从 Outlook 加载项调用 Web 服务](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) 。</span><span class="sxs-lookup"><span data-stu-id="23ad3-575">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="23ad3-576">不能通过 `makeEwsRequestAsync` 方法请求与“文件夹”关联的项。</span><span class="sxs-lookup"><span data-stu-id="23ad3-576">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="23ad3-577">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="23ad3-577">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="23ad3-p142">加载项必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。欲知使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定邮件加载项访问用户邮箱的权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="23ad3-580">注意：服务器管理员必须在 Client Access Server EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="23ad3-580">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="23ad3-581">版本差异</span><span class="sxs-lookup"><span data-stu-id="23ad3-581">Version differences</span></span>

<span data-ttu-id="23ad3-582">当你在较 15.0.4535.1004 版本更早的 Outlook 版本的邮件应用程序中使用 `makeEwsRequestAsync` 方法时，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="23ad3-582">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="23ad3-p143">当邮件应用在 Outlook 网页版中运行时，不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定邮件应用是正在 Outlook 中运行还是在 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的 Outlook 版本。</span><span class="sxs-lookup"><span data-stu-id="23ad3-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="23ad3-586">参数：</span><span class="sxs-lookup"><span data-stu-id="23ad3-586">Parameters:</span></span>

|<span data-ttu-id="23ad3-587">名称</span><span class="sxs-lookup"><span data-stu-id="23ad3-587">Name</span></span>| <span data-ttu-id="23ad3-588">类型</span><span class="sxs-lookup"><span data-stu-id="23ad3-588">Type</span></span>| <span data-ttu-id="23ad3-589">属性</span><span class="sxs-lookup"><span data-stu-id="23ad3-589">Attributes</span></span>| <span data-ttu-id="23ad3-590">说明</span><span class="sxs-lookup"><span data-stu-id="23ad3-590">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="23ad3-591">String</span><span class="sxs-lookup"><span data-stu-id="23ad3-591">String</span></span>||<span data-ttu-id="23ad3-592">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="23ad3-592">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="23ad3-593">function</span><span class="sxs-lookup"><span data-stu-id="23ad3-593">function</span></span>||<span data-ttu-id="23ad3-594">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="23ad3-594">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="23ad3-595">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="23ad3-595">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="23ad3-596">如果结果的大小超过 1 MB，将转而返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="23ad3-596">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="23ad3-597">Object</span><span class="sxs-lookup"><span data-stu-id="23ad3-597">Object</span></span>| <span data-ttu-id="23ad3-598">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="23ad3-598">&lt;optional&gt;</span></span>|<span data-ttu-id="23ad3-599">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="23ad3-599">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="23ad3-600">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-600">Requirements</span></span>

|<span data-ttu-id="23ad3-601">要求</span><span class="sxs-lookup"><span data-stu-id="23ad3-601">Requirement</span></span>| <span data-ttu-id="23ad3-602">值</span><span class="sxs-lookup"><span data-stu-id="23ad3-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="23ad3-603">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="23ad3-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="23ad3-604">1.0</span><span class="sxs-lookup"><span data-stu-id="23ad3-604">1.0</span></span>|
|[<span data-ttu-id="23ad3-605">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="23ad3-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="23ad3-606">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="23ad3-606">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="23ad3-607">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="23ad3-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="23ad3-608">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="23ad3-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="23ad3-609">示例</span><span class="sxs-lookup"><span data-stu-id="23ad3-609">Example</span></span>

<span data-ttu-id="23ad3-610">下面的示例调用 `makeEwsRequestAsync`  以使用  `GetItem` 操作获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="23ad3-610">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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