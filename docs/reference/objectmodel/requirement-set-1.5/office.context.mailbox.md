
# <a name="mailbox"></a><span data-ttu-id="c2708-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="c2708-101">mailbox</span></span>

### <span data-ttu-id="c2708-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="c2708-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="c2708-104">为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c2708-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c2708-105">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-105">Requirements</span></span>

|<span data-ttu-id="c2708-106">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-106">Requirement</span></span>| <span data-ttu-id="c2708-107">值</span><span class="sxs-lookup"><span data-stu-id="c2708-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-109">1.0</span></span>|
|[<span data-ttu-id="c2708-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-111">受限</span><span class="sxs-lookup"><span data-stu-id="c2708-111">Restricted</span></span>|
|[<span data-ttu-id="c2708-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="c2708-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c2708-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c2708-114">Members and methods</span></span>

| <span data-ttu-id="c2708-115">成员</span><span class="sxs-lookup"><span data-stu-id="c2708-115">Member</span></span> | <span data-ttu-id="c2708-116">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c2708-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="c2708-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="c2708-118">成员</span><span class="sxs-lookup"><span data-stu-id="c2708-118">Member</span></span> |
| [<span data-ttu-id="c2708-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="c2708-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="c2708-120">成员</span><span class="sxs-lookup"><span data-stu-id="c2708-120">Member</span></span> |
| [<span data-ttu-id="c2708-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c2708-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c2708-122">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-122">Method</span></span> |
| [<span data-ttu-id="c2708-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="c2708-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="c2708-124">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-124">Method</span></span> |
| [<span data-ttu-id="c2708-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c2708-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="c2708-126">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-126">Method</span></span> |
| [<span data-ttu-id="c2708-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="c2708-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="c2708-128">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-128">Method</span></span> |
| [<span data-ttu-id="c2708-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="c2708-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="c2708-130">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-130">Method</span></span> |
| [<span data-ttu-id="c2708-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c2708-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="c2708-132">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-132">Method</span></span> |
| [<span data-ttu-id="c2708-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="c2708-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="c2708-134">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-134">Method</span></span> |
| [<span data-ttu-id="c2708-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c2708-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="c2708-136">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-136">Method</span></span> |
| [<span data-ttu-id="c2708-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c2708-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="c2708-138">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-138">Method</span></span> |
| [<span data-ttu-id="c2708-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c2708-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="c2708-140">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-140">Method</span></span> |
| [<span data-ttu-id="c2708-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c2708-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="c2708-142">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-142">Method</span></span> |
| [<span data-ttu-id="c2708-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="c2708-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="c2708-144">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-144">Method</span></span> |
| [<span data-ttu-id="c2708-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c2708-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c2708-146">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c2708-147">命名空间</span><span class="sxs-lookup"><span data-stu-id="c2708-147">Namespaces</span></span>

<span data-ttu-id="c2708-148">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="c2708-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="c2708-149">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="c2708-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="c2708-150">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="c2708-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="c2708-151">成员</span><span class="sxs-lookup"><span data-stu-id="c2708-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="c2708-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="c2708-152">ewsUrl :String</span></span>

<span data-ttu-id="c2708-p102">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c2708-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-155">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="c2708-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c2708-p103">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="c2708-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c2708-158">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="c2708-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="c2708-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="c2708-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c2708-161">类型:</span><span class="sxs-lookup"><span data-stu-id="c2708-161">Type:</span></span>

*   <span data-ttu-id="c2708-162">String</span><span class="sxs-lookup"><span data-stu-id="c2708-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c2708-163">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-163">Requirements</span></span>

|<span data-ttu-id="c2708-164">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-164">Requirement</span></span>| <span data-ttu-id="c2708-165">值</span><span class="sxs-lookup"><span data-stu-id="c2708-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-166">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-167">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-167">1.0</span></span>|
|[<span data-ttu-id="c2708-168">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-169">ReadItem</span></span>|
|[<span data-ttu-id="c2708-170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="c2708-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="c2708-172">restUrl :String</span></span>

<span data-ttu-id="c2708-173">获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="c2708-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="c2708-174">`restUrl` 值可用于对用户邮箱进行 [REST API](https://docs.microsoft.com/outlook/rest/) 调用。</span><span class="sxs-lookup"><span data-stu-id="c2708-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="c2708-175">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `restUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="c2708-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="c2708-p105">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `restUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="c2708-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-178">连接到配置了自定义 REST URL 的 Exchange 2016 或更高版本本地安装的 Outlook 客户端将返回 `restUrl` 的无效值。</span><span class="sxs-lookup"><span data-stu-id="c2708-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="c2708-179">类型：</span><span class="sxs-lookup"><span data-stu-id="c2708-179">Type:</span></span>

*   <span data-ttu-id="c2708-180">String</span><span class="sxs-lookup"><span data-stu-id="c2708-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c2708-181">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-181">Requirements</span></span>

|<span data-ttu-id="c2708-182">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-182">Requirement</span></span>| <span data-ttu-id="c2708-183">值</span><span class="sxs-lookup"><span data-stu-id="c2708-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-185">1.5</span><span class="sxs-lookup"><span data-stu-id="c2708-185">1.5</span></span> |
|[<span data-ttu-id="c2708-186">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-186">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-187">ReadItem</span></span>|
|[<span data-ttu-id="c2708-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-188">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-189">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="c2708-190">方法</span><span class="sxs-lookup"><span data-stu-id="c2708-190">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c2708-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c2708-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c2708-192">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c2708-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c2708-193">目前，唯一受支持的事件类型是 `Office.EventType.ItemChanged`，用户选择一个新项目时将调用该类型。</span><span class="sxs-lookup"><span data-stu-id="c2708-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="c2708-194">此事件由实现可固定任务窗格的加载项使用，并允许加载项根据当前选定的项目刷新任务窗格 UI。</span><span class="sxs-lookup"><span data-stu-id="c2708-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-195">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-195">Parameters:</span></span>

| <span data-ttu-id="c2708-196">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-196">Name</span></span> | <span data-ttu-id="c2708-197">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-197">Type</span></span> | <span data-ttu-id="c2708-198">属性</span><span class="sxs-lookup"><span data-stu-id="c2708-198">Attributes</span></span> | <span data-ttu-id="c2708-199">说明</span><span class="sxs-lookup"><span data-stu-id="c2708-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c2708-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c2708-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c2708-201">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="c2708-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c2708-202">函数</span><span class="sxs-lookup"><span data-stu-id="c2708-202">Function</span></span> || <span data-ttu-id="c2708-p107">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="c2708-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c2708-206">Object</span><span class="sxs-lookup"><span data-stu-id="c2708-206">Object</span></span> | <span data-ttu-id="c2708-207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-207">&lt;optional&gt;</span></span> | <span data-ttu-id="c2708-208">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c2708-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c2708-209">对象</span><span class="sxs-lookup"><span data-stu-id="c2708-209">Object</span></span> | <span data-ttu-id="c2708-210">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-210">&lt;optional&gt;</span></span> | <span data-ttu-id="c2708-211">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c2708-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c2708-212">函数</span><span class="sxs-lookup"><span data-stu-id="c2708-212">function</span></span>| <span data-ttu-id="c2708-213">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-213">&lt;optional&gt;</span></span>|<span data-ttu-id="c2708-214">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c2708-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-215">Requirements</span><span class="sxs-lookup"><span data-stu-id="c2708-215">Requirements</span></span>

|<span data-ttu-id="c2708-216">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-216">Requirement</span></span>| <span data-ttu-id="c2708-217">值</span><span class="sxs-lookup"><span data-stu-id="c2708-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-219">1.5</span><span class="sxs-lookup"><span data-stu-id="c2708-219">1.5</span></span> |
|[<span data-ttu-id="c2708-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-221">ReadItem</span></span> |
|[<span data-ttu-id="c2708-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-223">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-223">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2708-224">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-224">Example</span></span>

```js
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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="c2708-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c2708-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c2708-226">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="c2708-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-227">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c2708-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c2708-p108">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="c2708-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-230">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-230">Parameters:</span></span>

|<span data-ttu-id="c2708-231">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-231">Name</span></span>| <span data-ttu-id="c2708-232">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-232">Type</span></span>| <span data-ttu-id="c2708-233">描述</span><span class="sxs-lookup"><span data-stu-id="c2708-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c2708-234">字符串</span><span class="sxs-lookup"><span data-stu-id="c2708-234">String</span></span>|<span data-ttu-id="c2708-235">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="c2708-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="c2708-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c2708-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="c2708-237">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="c2708-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-238">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-238">Requirements</span></span>

|<span data-ttu-id="c2708-239">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-239">Requirement</span></span>| <span data-ttu-id="c2708-240">值</span><span class="sxs-lookup"><span data-stu-id="c2708-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-242">1.3</span><span class="sxs-lookup"><span data-stu-id="c2708-242">1.3</span></span>|
|[<span data-ttu-id="c2708-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-244">受限</span><span class="sxs-lookup"><span data-stu-id="c2708-244">Restricted</span></span>|
|[<span data-ttu-id="c2708-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-246">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c2708-247">返回：</span><span class="sxs-lookup"><span data-stu-id="c2708-247">Returns:</span></span>

<span data-ttu-id="c2708-248">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="c2708-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c2708-249">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-249">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="c2708-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="c2708-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="c2708-251">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="c2708-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="c2708-p109">Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="c2708-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="c2708-p110">如果邮件应用程序在 Outlook 中运行，`convertToLocalClientTime` 方法将返回一个值设置为客户端计算机时区的字典对象。如果邮件应用程序在 Outlook Web App 中运行，`convertToLocalClientTime` 方法将返回值设置为 EAC 中指定的时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="c2708-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-257">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-257">Parameters:</span></span>

|<span data-ttu-id="c2708-258">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-258">Name</span></span>| <span data-ttu-id="c2708-259">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-259">Type</span></span>| <span data-ttu-id="c2708-260">描述</span><span class="sxs-lookup"><span data-stu-id="c2708-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="c2708-261">日期</span><span class="sxs-lookup"><span data-stu-id="c2708-261">Date</span></span>|<span data-ttu-id="c2708-262">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="c2708-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-263">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-263">Requirements</span></span>

|<span data-ttu-id="c2708-264">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-264">Requirement</span></span>| <span data-ttu-id="c2708-265">值</span><span class="sxs-lookup"><span data-stu-id="c2708-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-266">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-267">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-267">1.0</span></span>|
|[<span data-ttu-id="c2708-268">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-269">ReadItem</span></span>|
|[<span data-ttu-id="c2708-270">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-271">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-271">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c2708-272">返回：</span><span class="sxs-lookup"><span data-stu-id="c2708-272">Returns:</span></span>

<span data-ttu-id="c2708-273">类型：[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="c2708-273">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="c2708-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c2708-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c2708-275">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="c2708-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-276">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c2708-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c2708-p111">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="c2708-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-279">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-279">Parameters:</span></span>

|<span data-ttu-id="c2708-280">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-280">Name</span></span>| <span data-ttu-id="c2708-281">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-281">Type</span></span>| <span data-ttu-id="c2708-282">描述</span><span class="sxs-lookup"><span data-stu-id="c2708-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c2708-283">字符串</span><span class="sxs-lookup"><span data-stu-id="c2708-283">String</span></span>|<span data-ttu-id="c2708-284">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="c2708-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="c2708-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c2708-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="c2708-286">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="c2708-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-287">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-287">Requirements</span></span>

|<span data-ttu-id="c2708-288">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-288">Requirement</span></span>| <span data-ttu-id="c2708-289">值</span><span class="sxs-lookup"><span data-stu-id="c2708-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-290">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-291">1.3</span><span class="sxs-lookup"><span data-stu-id="c2708-291">1.3</span></span>|
|[<span data-ttu-id="c2708-292">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-292">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-293">受限</span><span class="sxs-lookup"><span data-stu-id="c2708-293">Restricted</span></span>|
|[<span data-ttu-id="c2708-294">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-294">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-295">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-295">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c2708-296">返回：</span><span class="sxs-lookup"><span data-stu-id="c2708-296">Returns:</span></span>

<span data-ttu-id="c2708-297">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="c2708-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c2708-298">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-298">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="c2708-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="c2708-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="c2708-300">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="c2708-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="c2708-301">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="c2708-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-302">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-302">Parameters:</span></span>

|<span data-ttu-id="c2708-303">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-303">Name</span></span>| <span data-ttu-id="c2708-304">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-304">Type</span></span>| <span data-ttu-id="c2708-305">说明</span><span class="sxs-lookup"><span data-stu-id="c2708-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="c2708-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c2708-306">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="c2708-307">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="c2708-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-308">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-308">Requirements</span></span>

|<span data-ttu-id="c2708-309">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-309">Requirement</span></span>| <span data-ttu-id="c2708-310">值</span><span class="sxs-lookup"><span data-stu-id="c2708-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-312">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-312">1.0</span></span>|
|[<span data-ttu-id="c2708-313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-314">ReadItem</span></span>|
|[<span data-ttu-id="c2708-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-316">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-316">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c2708-317">返回：</span><span class="sxs-lookup"><span data-stu-id="c2708-317">Returns:</span></span>

<span data-ttu-id="c2708-318">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="c2708-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="c2708-319">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="c2708-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c2708-320">日期</span><span class="sxs-lookup"><span data-stu-id="c2708-320">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="c2708-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c2708-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="c2708-322">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="c2708-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-323">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c2708-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c2708-324">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="c2708-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c2708-p112">在 Outlook for Mac 中，您可以使用此方法来显示不属于定期系列的单个约会，或显示定期系列的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。</span><span class="sxs-lookup"><span data-stu-id="c2708-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="c2708-327">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="c2708-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="c2708-328">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="c2708-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-329">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-329">Parameters:</span></span>

|<span data-ttu-id="c2708-330">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-330">Name</span></span>| <span data-ttu-id="c2708-331">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-331">Type</span></span>| <span data-ttu-id="c2708-332">描述</span><span class="sxs-lookup"><span data-stu-id="c2708-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c2708-333">字符串</span><span class="sxs-lookup"><span data-stu-id="c2708-333">String</span></span>|<span data-ttu-id="c2708-334">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="c2708-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-335">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-335">Requirements</span></span>

|<span data-ttu-id="c2708-336">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-336">Requirement</span></span>| <span data-ttu-id="c2708-337">值</span><span class="sxs-lookup"><span data-stu-id="c2708-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-338">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-339">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-339">1.0</span></span>|
|[<span data-ttu-id="c2708-340">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-340">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-341">ReadItem</span></span>|
|[<span data-ttu-id="c2708-342">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-342">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-343">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-343">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2708-344">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="c2708-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c2708-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="c2708-346">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="c2708-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-347">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c2708-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c2708-348">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="c2708-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c2708-349">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="c2708-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="c2708-350">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="c2708-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="c2708-p113">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="c2708-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-353">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-353">Parameters:</span></span>

|<span data-ttu-id="c2708-354">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-354">Name</span></span>| <span data-ttu-id="c2708-355">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-355">Type</span></span>| <span data-ttu-id="c2708-356">描述</span><span class="sxs-lookup"><span data-stu-id="c2708-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c2708-357">字符串</span><span class="sxs-lookup"><span data-stu-id="c2708-357">String</span></span>|<span data-ttu-id="c2708-358">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="c2708-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-359">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-359">Requirements</span></span>

|<span data-ttu-id="c2708-360">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-360">Requirement</span></span>| <span data-ttu-id="c2708-361">值</span><span class="sxs-lookup"><span data-stu-id="c2708-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-363">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-363">1.0</span></span>|
|[<span data-ttu-id="c2708-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-365">ReadItem</span></span>|
|[<span data-ttu-id="c2708-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-367">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2708-368">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="c2708-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="c2708-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="c2708-370">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="c2708-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-371">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c2708-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c2708-p114">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="c2708-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c2708-p115">在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="c2708-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="c2708-p116">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="c2708-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="c2708-379">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="c2708-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-380">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-380">Parameters:</span></span>

|<span data-ttu-id="c2708-381">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-381">Name</span></span>| <span data-ttu-id="c2708-382">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-382">Type</span></span>| <span data-ttu-id="c2708-383">描述</span><span class="sxs-lookup"><span data-stu-id="c2708-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c2708-384">对象</span><span class="sxs-lookup"><span data-stu-id="c2708-384">Object</span></span> | <span data-ttu-id="c2708-385">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="c2708-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="c2708-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c2708-p117">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="c2708-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="c2708-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c2708-p118">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="c2708-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="c2708-392">日期</span><span class="sxs-lookup"><span data-stu-id="c2708-392">Date</span></span> | <span data-ttu-id="c2708-393">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c2708-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="c2708-394">Date</span><span class="sxs-lookup"><span data-stu-id="c2708-394">Date</span></span> | <span data-ttu-id="c2708-395">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c2708-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="c2708-396">String</span><span class="sxs-lookup"><span data-stu-id="c2708-396">String</span></span> | <span data-ttu-id="c2708-p119">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c2708-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="c2708-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="c2708-p120">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="c2708-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c2708-402">String</span><span class="sxs-lookup"><span data-stu-id="c2708-402">String</span></span> | <span data-ttu-id="c2708-p121">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c2708-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="c2708-405">字符串</span><span class="sxs-lookup"><span data-stu-id="c2708-405">String</span></span> | <span data-ttu-id="c2708-p122">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c2708-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c2708-408">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-408">Requirements</span></span>

|<span data-ttu-id="c2708-409">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-409">Requirement</span></span>| <span data-ttu-id="c2708-410">值</span><span class="sxs-lookup"><span data-stu-id="c2708-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-412">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-412">1.0</span></span>|
|[<span data-ttu-id="c2708-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-414">ReadItem</span></span>|
|[<span data-ttu-id="c2708-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-416">阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2708-417">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="c2708-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c2708-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="c2708-419">获取一个包含用于调用 REST API 或 Exchange Web 服务的令牌的字符串。</span><span class="sxs-lookup"><span data-stu-id="c2708-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="c2708-p123">
            `getCallbackTokenAsync\` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="c2708-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-422">建议加载项尽可能地使用 REST API 而不是 Exchange Web 服务。</span><span class="sxs-lookup"><span data-stu-id="c2708-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="c2708-423">**REST 令牌**</span><span class="sxs-lookup"><span data-stu-id="c2708-423">**REST Tokens**</span></span>

<span data-ttu-id="c2708-p124">请求 REST 令牌时 (`options.isRest = true`) 时，生成的令牌将无法对 Exchange Web 服务调用进行身份验证。令牌的作用域限制为对当前项及其附件的只读访问，除非外接程序在其清单中指定了 [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) 权限。如果指定了 `ReadWriteMailbox` 权限，则生成的令牌将授予对邮件、日历和联系人的读/写权限，包括发送邮件的功能。</span><span class="sxs-lookup"><span data-stu-id="c2708-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="c2708-427">在进行 REST API 调用时，外接程序应使用 `restUrl` 属性来确定要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="c2708-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="c2708-428">**EWS 令牌**</span><span class="sxs-lookup"><span data-stu-id="c2708-428">**EWS Tokens**</span></span>

<span data-ttu-id="c2708-p125">请求 EWS 令牌 (`options.isRest = false`) 时，生成的令牌将无法对 REST API 调用进行身份验证。令牌的作用域限制为访问当前项。</span><span class="sxs-lookup"><span data-stu-id="c2708-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="c2708-431">外接程序应使用 `ewsUrl` 属性来确定进行 EWS 调用时要使用的正确 URL。</span><span class="sxs-lookup"><span data-stu-id="c2708-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-432">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-432">Parameters:</span></span>

|<span data-ttu-id="c2708-433">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-433">Name</span></span>| <span data-ttu-id="c2708-434">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-434">Type</span></span>| <span data-ttu-id="c2708-435">属性</span><span class="sxs-lookup"><span data-stu-id="c2708-435">Attributes</span></span>| <span data-ttu-id="c2708-436">说明</span><span class="sxs-lookup"><span data-stu-id="c2708-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="c2708-437">Object</span><span class="sxs-lookup"><span data-stu-id="c2708-437">Object</span></span> | <span data-ttu-id="c2708-438">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-438">&lt;optional&gt;</span></span> | <span data-ttu-id="c2708-439">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c2708-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="c2708-440">布尔值</span><span class="sxs-lookup"><span data-stu-id="c2708-440">Boolean</span></span> |  <span data-ttu-id="c2708-441">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-441">&lt;optional&gt;</span></span> | <span data-ttu-id="c2708-p126">确定所提供的令牌是否将用于 Outlook REST API 或 Exchange Web 服务。默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="c2708-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c2708-444">Object</span><span class="sxs-lookup"><span data-stu-id="c2708-444">Object</span></span> |  <span data-ttu-id="c2708-445">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-445">&lt;optional&gt;</span></span> | <span data-ttu-id="c2708-446">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="c2708-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="c2708-447">函数</span><span class="sxs-lookup"><span data-stu-id="c2708-447">function</span></span>||<span data-ttu-id="c2708-p127">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="c2708-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-450">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-450">Requirements</span></span>

|<span data-ttu-id="c2708-451">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-451">Requirement</span></span>| <span data-ttu-id="c2708-452">值</span><span class="sxs-lookup"><span data-stu-id="c2708-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-453">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-454">1.5</span><span class="sxs-lookup"><span data-stu-id="c2708-454">1.5</span></span> |
|[<span data-ttu-id="c2708-455">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-455">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-456">ReadItem</span></span>|
|[<span data-ttu-id="c2708-457">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-457">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-458">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-458">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2708-459">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-459">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="c2708-460">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c2708-460">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c2708-461">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="c2708-461">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="c2708-p128">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="c2708-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="c2708-p129">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="c2708-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c2708-467">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="c2708-467">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="c2708-p130">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="c2708-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-470">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-470">Parameters:</span></span>

|<span data-ttu-id="c2708-471">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-471">Name</span></span>| <span data-ttu-id="c2708-472">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-472">Type</span></span>| <span data-ttu-id="c2708-473">属性</span><span class="sxs-lookup"><span data-stu-id="c2708-473">Attributes</span></span>| <span data-ttu-id="c2708-474">说明</span><span class="sxs-lookup"><span data-stu-id="c2708-474">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c2708-475">函数</span><span class="sxs-lookup"><span data-stu-id="c2708-475">function</span></span>||<span data-ttu-id="c2708-p131">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="c2708-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="c2708-478">对象</span><span class="sxs-lookup"><span data-stu-id="c2708-478">Object</span></span>| <span data-ttu-id="c2708-479">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-479">&lt;optional&gt;</span></span>|<span data-ttu-id="c2708-480">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="c2708-480">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-481">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-481">Requirements</span></span>

|<span data-ttu-id="c2708-482">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-482">Requirement</span></span>| <span data-ttu-id="c2708-483">值</span><span class="sxs-lookup"><span data-stu-id="c2708-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-485">1.3</span><span class="sxs-lookup"><span data-stu-id="c2708-485">1.3</span></span>|
|[<span data-ttu-id="c2708-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-487">ReadItem</span></span>|
|[<span data-ttu-id="c2708-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-489">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-489">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2708-490">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-490">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="c2708-491">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c2708-491">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c2708-492">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="c2708-492">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="c2708-493">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](https://docs.microsoft.com/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="c2708-493">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-494">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-494">Parameters:</span></span>

|<span data-ttu-id="c2708-495">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-495">Name</span></span>| <span data-ttu-id="c2708-496">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-496">Type</span></span>| <span data-ttu-id="c2708-497">属性</span><span class="sxs-lookup"><span data-stu-id="c2708-497">Attributes</span></span>| <span data-ttu-id="c2708-498">说明</span><span class="sxs-lookup"><span data-stu-id="c2708-498">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c2708-499">函数</span><span class="sxs-lookup"><span data-stu-id="c2708-499">function</span></span>||<span data-ttu-id="c2708-500">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c2708-500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c2708-501">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="c2708-501">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="c2708-502">对象</span><span class="sxs-lookup"><span data-stu-id="c2708-502">Object</span></span>| <span data-ttu-id="c2708-503">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-503">&lt;optional&gt;</span></span>|<span data-ttu-id="c2708-504">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="c2708-504">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-505">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-505">Requirements</span></span>

|<span data-ttu-id="c2708-506">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-506">Requirement</span></span>| <span data-ttu-id="c2708-507">值</span><span class="sxs-lookup"><span data-stu-id="c2708-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-509">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-509">1.0</span></span>|
|[<span data-ttu-id="c2708-510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-511">ReadItem</span></span>|
|[<span data-ttu-id="c2708-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-513">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2708-514">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-514">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="c2708-515">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c2708-515">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="c2708-516">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="c2708-516">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-517">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="c2708-517">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="c2708-518">在 Outlook for iOS 或 Outlook for Android 中</span><span class="sxs-lookup"><span data-stu-id="c2708-518">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="c2708-519">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="c2708-519">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="c2708-520">在这些情况下，加载项应该[使用 REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="c2708-520">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="c2708-521">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="c2708-521">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="c2708-522">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="c2708-522">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="c2708-523">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="c2708-523">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="c2708-524">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="c2708-524">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="c2708-p133">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="c2708-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="c2708-527">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="c2708-527">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="c2708-528">版本差异</span><span class="sxs-lookup"><span data-stu-id="c2708-528">Version differences</span></span>

<span data-ttu-id="c2708-529">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="c2708-529">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="c2708-p134">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="c2708-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-533">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-533">Parameters:</span></span>

|<span data-ttu-id="c2708-534">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-534">Name</span></span>| <span data-ttu-id="c2708-535">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-535">Type</span></span>| <span data-ttu-id="c2708-536">属性</span><span class="sxs-lookup"><span data-stu-id="c2708-536">Attributes</span></span>| <span data-ttu-id="c2708-537">说明</span><span class="sxs-lookup"><span data-stu-id="c2708-537">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c2708-538">字符串</span><span class="sxs-lookup"><span data-stu-id="c2708-538">String</span></span>||<span data-ttu-id="c2708-539">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="c2708-539">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="c2708-540">函数</span><span class="sxs-lookup"><span data-stu-id="c2708-540">function</span></span>||<span data-ttu-id="c2708-541">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c2708-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c2708-542">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="c2708-542">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="c2708-543">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="c2708-543">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="c2708-544">对象</span><span class="sxs-lookup"><span data-stu-id="c2708-544">Object</span></span>| <span data-ttu-id="c2708-545">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-545">&lt;optional&gt;</span></span>|<span data-ttu-id="c2708-546">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="c2708-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-547">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-547">Requirements</span></span>

|<span data-ttu-id="c2708-548">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-548">Requirement</span></span>| <span data-ttu-id="c2708-549">值</span><span class="sxs-lookup"><span data-stu-id="c2708-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-550">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-551">1.0</span><span class="sxs-lookup"><span data-stu-id="c2708-551">1.0</span></span>|
|[<span data-ttu-id="c2708-552">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-553">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="c2708-553">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="c2708-554">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-555">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-555">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c2708-556">示例</span><span class="sxs-lookup"><span data-stu-id="c2708-556">Example</span></span>

<span data-ttu-id="c2708-557">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="c2708-557">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c2708-558">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c2708-558">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c2708-559">删除支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c2708-559">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="c2708-560">当前，唯一支持的事件类型是 `Office.EventType.ItemChanged`。</span><span class="sxs-lookup"><span data-stu-id="c2708-560">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c2708-561">参数：</span><span class="sxs-lookup"><span data-stu-id="c2708-561">Parameters:</span></span>

| <span data-ttu-id="c2708-562">名称</span><span class="sxs-lookup"><span data-stu-id="c2708-562">Name</span></span> | <span data-ttu-id="c2708-563">类型</span><span class="sxs-lookup"><span data-stu-id="c2708-563">Type</span></span> | <span data-ttu-id="c2708-564">属性</span><span class="sxs-lookup"><span data-stu-id="c2708-564">Attributes</span></span> | <span data-ttu-id="c2708-565">说明</span><span class="sxs-lookup"><span data-stu-id="c2708-565">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c2708-566">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c2708-566">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c2708-567">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="c2708-567">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c2708-568">函数</span><span class="sxs-lookup"><span data-stu-id="c2708-568">Function</span></span> || <span data-ttu-id="c2708-p136">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="c2708-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c2708-572">Object</span><span class="sxs-lookup"><span data-stu-id="c2708-572">Object</span></span> | <span data-ttu-id="c2708-573">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-573">&lt;optional&gt;</span></span> | <span data-ttu-id="c2708-574">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c2708-574">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c2708-575">对象</span><span class="sxs-lookup"><span data-stu-id="c2708-575">Object</span></span> | <span data-ttu-id="c2708-576">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-576">&lt;optional&gt;</span></span> | <span data-ttu-id="c2708-577">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c2708-577">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c2708-578">函数</span><span class="sxs-lookup"><span data-stu-id="c2708-578">function</span></span>| <span data-ttu-id="c2708-579">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c2708-579">&lt;optional&gt;</span></span>|<span data-ttu-id="c2708-580">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c2708-580">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c2708-581">Requirements</span><span class="sxs-lookup"><span data-stu-id="c2708-581">Requirements</span></span>

|<span data-ttu-id="c2708-582">要求</span><span class="sxs-lookup"><span data-stu-id="c2708-582">Requirement</span></span>| <span data-ttu-id="c2708-583">值</span><span class="sxs-lookup"><span data-stu-id="c2708-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="c2708-584">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c2708-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c2708-585">1.5</span><span class="sxs-lookup"><span data-stu-id="c2708-585">1.5</span></span> |
|[<span data-ttu-id="c2708-586">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c2708-586">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c2708-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c2708-587">ReadItem</span></span> |
|[<span data-ttu-id="c2708-588">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c2708-588">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c2708-589">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c2708-589">Compose or read</span></span>|