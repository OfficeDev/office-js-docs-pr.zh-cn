
# <a name="mailbox"></a><span data-ttu-id="458a4-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="458a4-101">mailbox</span></span>

### <span data-ttu-id="458a4-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="458a4-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="458a4-104">为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="458a4-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="458a4-105">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-105">Requirements</span></span>

|<span data-ttu-id="458a4-106">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-106">Requirement</span></span>| <span data-ttu-id="458a4-107">值</span><span class="sxs-lookup"><span data-stu-id="458a4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-109">1.0</span></span>|
|[<span data-ttu-id="458a4-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-111">受限</span><span class="sxs-lookup"><span data-stu-id="458a4-111">Restricted</span></span>|
|[<span data-ttu-id="458a4-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="458a4-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="458a4-114">命名空间</span><span class="sxs-lookup"><span data-stu-id="458a4-114">Namespaces</span></span>

<span data-ttu-id="458a4-115">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="458a4-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="458a4-116">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 外接程序中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="458a4-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="458a4-117">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 外接程序中的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="458a4-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="458a4-118">成员</span><span class="sxs-lookup"><span data-stu-id="458a4-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="458a4-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="458a4-119">ewsUrl :String</span></span>

<span data-ttu-id="458a4-p102">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 终结点的 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="458a4-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="458a4-122">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="458a4-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="458a4-p103">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="458a4-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="458a4-125">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `ewsUrl` 成员。</span><span class="sxs-lookup"><span data-stu-id="458a4-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="458a4-p104">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法，才能使用 `ewsUrl` 成员。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="458a4-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="458a4-128">类型:</span><span class="sxs-lookup"><span data-stu-id="458a4-128">Type:</span></span>

*   <span data-ttu-id="458a4-129">String</span><span class="sxs-lookup"><span data-stu-id="458a4-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="458a4-130">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-130">Requirements</span></span>

|<span data-ttu-id="458a4-131">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-131">Requirement</span></span>| <span data-ttu-id="458a4-132">值</span><span class="sxs-lookup"><span data-stu-id="458a4-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-133">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-134">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-134">1.0</span></span>|
|[<span data-ttu-id="458a4-135">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="458a4-136">ReadItem</span></span>|
|[<span data-ttu-id="458a4-137">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-138">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-138">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="458a4-139">方法</span><span class="sxs-lookup"><span data-stu-id="458a4-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="458a4-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="458a4-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="458a4-141">将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="458a4-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="458a4-142">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="458a4-142">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="458a4-p105">通过 REST API 检索的项 ID（如 [Outlook 邮件 API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）使用与 Exchange Web 服务 (EWS) 所使用格式不同的格式。`convertToEwsId` 方法将 REST 格式化的 ID 转换为正确的 EWS 格式。</span><span class="sxs-lookup"><span data-stu-id="458a4-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-145">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-145">Parameters:</span></span>

|<span data-ttu-id="458a4-146">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-146">Name</span></span>| <span data-ttu-id="458a4-147">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-147">Type</span></span>| <span data-ttu-id="458a4-148">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="458a4-149">字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-149">String</span></span>|<span data-ttu-id="458a4-150">Outlook REST API 的格式化的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="458a4-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="458a4-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="458a4-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="458a4-152">指示用于检索项目 ID 的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="458a4-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-153">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-153">Requirements</span></span>

|<span data-ttu-id="458a4-154">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-154">Requirement</span></span>| <span data-ttu-id="458a4-155">值</span><span class="sxs-lookup"><span data-stu-id="458a4-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-156">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-157">1.3</span><span class="sxs-lookup"><span data-stu-id="458a4-157">1.3</span></span>|
|[<span data-ttu-id="458a4-158">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-159">受限</span><span class="sxs-lookup"><span data-stu-id="458a4-159">Restricted</span></span>|
|[<span data-ttu-id="458a4-160">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-161">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-161">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="458a4-162">返回：</span><span class="sxs-lookup"><span data-stu-id="458a4-162">Returns:</span></span>

<span data-ttu-id="458a4-163">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="458a4-164">示例</span><span class="sxs-lookup"><span data-stu-id="458a4-164">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a><span data-ttu-id="458a4-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="458a4-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span></span>

<span data-ttu-id="458a4-166">获取包含以本地客户端时间表示的时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="458a4-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="458a4-p106">Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="458a4-p106">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="458a4-p107">如果邮件应用程序在 Outlook 中运行，`convertToLocalClientTime` 方法将返回一个值设置为客户端计算机时区的字典对象。如果邮件应用程序在 Outlook Web App 中运行，`convertToLocalClientTime` 方法将返回值设置为 EAC 中指定的时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="458a4-p107">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-172">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-172">Parameters:</span></span>

|<span data-ttu-id="458a4-173">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-173">Name</span></span>| <span data-ttu-id="458a4-174">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-174">Type</span></span>| <span data-ttu-id="458a4-175">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="458a4-176">日期</span><span class="sxs-lookup"><span data-stu-id="458a4-176">Date</span></span>|<span data-ttu-id="458a4-177">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="458a4-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-178">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-178">Requirements</span></span>

|<span data-ttu-id="458a4-179">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-179">Requirement</span></span>| <span data-ttu-id="458a4-180">值</span><span class="sxs-lookup"><span data-stu-id="458a4-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-182">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-182">1.0</span></span>|
|[<span data-ttu-id="458a4-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="458a4-184">ReadItem</span></span>|
|[<span data-ttu-id="458a4-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-186">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="458a4-187">返回：</span><span class="sxs-lookup"><span data-stu-id="458a4-187">Returns:</span></span>

<span data-ttu-id="458a4-188">类型：[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="458a4-188">Type: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="458a4-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="458a4-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="458a4-190">将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="458a4-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="458a4-191">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="458a4-191">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="458a4-p108">与 REST API 所使用的格式比较，通过 EWS 或通过 `itemId` 属性检索的项目 ID 使用不同的格式（例如 [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) 或 [Microsoft Graph](https://graph.microsoft.io/)）。`convertToRestId` 方法将 EWS 格式化的 ID 转换为正确的 REST 格式。</span><span class="sxs-lookup"><span data-stu-id="458a4-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-194">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-194">Parameters:</span></span>

|<span data-ttu-id="458a4-195">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-195">Name</span></span>| <span data-ttu-id="458a4-196">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-196">Type</span></span>| <span data-ttu-id="458a4-197">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="458a4-198">字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-198">String</span></span>|<span data-ttu-id="458a4-199">适用于 Exchange Web 服务 (EWS) 的项目 ID 格式化。</span><span class="sxs-lookup"><span data-stu-id="458a4-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="458a4-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="458a4-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="458a4-201">值指示转换的 ID 所使用的 Outlook REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="458a4-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-202">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-202">Requirements</span></span>

|<span data-ttu-id="458a4-203">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-203">Requirement</span></span>| <span data-ttu-id="458a4-204">值</span><span class="sxs-lookup"><span data-stu-id="458a4-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-205">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-206">1.3</span><span class="sxs-lookup"><span data-stu-id="458a4-206">1.3</span></span>|
|[<span data-ttu-id="458a4-207">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-207">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-208">受限</span><span class="sxs-lookup"><span data-stu-id="458a4-208">Restricted</span></span>|
|[<span data-ttu-id="458a4-209">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-210">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-210">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="458a4-211">返回：</span><span class="sxs-lookup"><span data-stu-id="458a4-211">Returns:</span></span>

<span data-ttu-id="458a4-212">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="458a4-213">示例</span><span class="sxs-lookup"><span data-stu-id="458a4-213">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="458a4-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="458a4-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="458a4-215">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="458a4-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="458a4-216">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="458a4-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-217">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-217">Parameters:</span></span>

|<span data-ttu-id="458a4-218">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-218">Name</span></span>| <span data-ttu-id="458a4-219">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-219">Type</span></span>| <span data-ttu-id="458a4-220">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="458a4-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="458a4-221">LocalClientTime</span></span>](/javascript/api/outlook_1_3/office.LocalClientTime)|<span data-ttu-id="458a4-222">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="458a4-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-223">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-223">Requirements</span></span>

|<span data-ttu-id="458a4-224">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-224">Requirement</span></span>| <span data-ttu-id="458a4-225">值</span><span class="sxs-lookup"><span data-stu-id="458a4-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-226">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-227">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-227">1.0</span></span>|
|[<span data-ttu-id="458a4-228">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="458a4-229">ReadItem</span></span>|
|[<span data-ttu-id="458a4-230">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-231">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-231">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="458a4-232">返回：</span><span class="sxs-lookup"><span data-stu-id="458a4-232">Returns:</span></span>

<span data-ttu-id="458a4-233">包含以 UTC 表示的时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="458a4-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="458a4-234">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="458a4-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="458a4-235">日期</span><span class="sxs-lookup"><span data-stu-id="458a4-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="458a4-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="458a4-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="458a4-237">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="458a4-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="458a4-238">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="458a4-238">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="458a4-239">`displayAppointmentForm` 方法将打开桌面新窗口中或移动设备对话框中的现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="458a4-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="458a4-p109">在 Outlook for Mac 中，您可以使用此方法来显示不属于定期系列的单个约会，或显示定期系列的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。</span><span class="sxs-lookup"><span data-stu-id="458a4-p109">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="458a4-242">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="458a4-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="458a4-243">如果指定的项标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="458a4-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-244">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-244">Parameters:</span></span>

|<span data-ttu-id="458a4-245">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-245">Name</span></span>| <span data-ttu-id="458a4-246">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-246">Type</span></span>| <span data-ttu-id="458a4-247">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="458a4-248">字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-248">String</span></span>|<span data-ttu-id="458a4-249">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="458a4-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-250">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-250">Requirements</span></span>

|<span data-ttu-id="458a4-251">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-251">Requirement</span></span>| <span data-ttu-id="458a4-252">值</span><span class="sxs-lookup"><span data-stu-id="458a4-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-253">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-254">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-254">1.0</span></span>|
|[<span data-ttu-id="458a4-255">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="458a4-256">ReadItem</span></span>|
|[<span data-ttu-id="458a4-257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-258">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-258">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="458a4-259">示例</span><span class="sxs-lookup"><span data-stu-id="458a4-259">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="458a4-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="458a4-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="458a4-261">显示现有邮件。</span><span class="sxs-lookup"><span data-stu-id="458a4-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="458a4-262">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="458a4-262">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="458a4-263">`displayMessageForm` 方法将打开桌面新窗口中或移动设备对话框中的现有邮件。</span><span class="sxs-lookup"><span data-stu-id="458a4-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="458a4-264">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="458a4-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="458a4-265">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="458a4-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="458a4-p110">不要使用包含表示约会的 `itemId` 的 `displayMessageForm`。使用 `displayAppointmentForm` 方法显示现有的约会，并使用 `displayNewAppointmentForm` 显示窗体以新建约会。</span><span class="sxs-lookup"><span data-stu-id="458a4-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-268">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-268">Parameters:</span></span>

|<span data-ttu-id="458a4-269">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-269">Name</span></span>| <span data-ttu-id="458a4-270">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-270">Type</span></span>| <span data-ttu-id="458a4-271">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="458a4-272">字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-272">String</span></span>|<span data-ttu-id="458a4-273">现有消息的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="458a4-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-274">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-274">Requirements</span></span>

|<span data-ttu-id="458a4-275">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-275">Requirement</span></span>| <span data-ttu-id="458a4-276">值</span><span class="sxs-lookup"><span data-stu-id="458a4-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-277">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-278">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-278">1.0</span></span>|
|[<span data-ttu-id="458a4-279">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="458a4-280">ReadItem</span></span>|
|[<span data-ttu-id="458a4-281">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-282">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-282">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="458a4-283">示例</span><span class="sxs-lookup"><span data-stu-id="458a4-283">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="458a4-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="458a4-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="458a4-285">显示用于新建日历约会的表单。</span><span class="sxs-lookup"><span data-stu-id="458a4-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="458a4-286">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="458a4-286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="458a4-p111">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充约会窗体字段。</span><span class="sxs-lookup"><span data-stu-id="458a4-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="458a4-p112">在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含“**保存**”按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="458a4-p112">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="458a4-p113">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="458a4-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="458a4-294">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="458a4-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-295">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-295">Parameters:</span></span>

|<span data-ttu-id="458a4-296">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-296">Name</span></span>| <span data-ttu-id="458a4-297">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-297">Type</span></span>| <span data-ttu-id="458a4-298">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="458a4-299">对象</span><span class="sxs-lookup"><span data-stu-id="458a4-299">Object</span></span> | <span data-ttu-id="458a4-300">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="458a4-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="458a4-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="458a4-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="458a4-p114">包含电子邮件地址的字符串数组或包含约会的每个必需与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="458a4-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="458a4-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="458a4-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="458a4-p115">包含电子邮件地址的字符串数组或包含约会的每个可选与会者的 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="458a4-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="458a4-307">日期</span><span class="sxs-lookup"><span data-stu-id="458a4-307">Date</span></span> | <span data-ttu-id="458a4-308">指定约会的开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="458a4-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="458a4-309">Date</span><span class="sxs-lookup"><span data-stu-id="458a4-309">Date</span></span> | <span data-ttu-id="458a4-310">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="458a4-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="458a4-311">String</span><span class="sxs-lookup"><span data-stu-id="458a4-311">String</span></span> | <span data-ttu-id="458a4-p116">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="458a4-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="458a4-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="458a4-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="458a4-p117">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="458a4-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="458a4-317">字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-317">String</span></span> | <span data-ttu-id="458a4-p118">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="458a4-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="458a4-320">字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-320">String</span></span> | <span data-ttu-id="458a4-p119">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="458a4-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="458a4-323">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-323">Requirements</span></span>

|<span data-ttu-id="458a4-324">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-324">Requirement</span></span>| <span data-ttu-id="458a4-325">值</span><span class="sxs-lookup"><span data-stu-id="458a4-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-326">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-327">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-327">1.0</span></span>|
|[<span data-ttu-id="458a4-328">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="458a4-329">ReadItem</span></span>|
|[<span data-ttu-id="458a4-330">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-331">阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="458a4-332">示例</span><span class="sxs-lookup"><span data-stu-id="458a4-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="458a4-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="458a4-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="458a4-334">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="458a4-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="458a4-p120">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="458a4-p120">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="458a4-p121">可以将令牌和附件标识符或项标识符传递到第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="458a4-p121">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="458a4-340">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用阅读模式中的 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="458a4-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="458a4-p122">在撰写模式中，必须调用 [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) 方法来获取传递给 `getCallbackTokenAsync` 方法的项目标识符。应用必须具有调用 `saveAsync` 方法的 **ReadWriteItem** 权限。</span><span class="sxs-lookup"><span data-stu-id="458a4-p122">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-343">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-343">Parameters:</span></span>

|<span data-ttu-id="458a4-344">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-344">Name</span></span>| <span data-ttu-id="458a4-345">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-345">Type</span></span>| <span data-ttu-id="458a4-346">属性</span><span class="sxs-lookup"><span data-stu-id="458a4-346">Attributes</span></span>| <span data-ttu-id="458a4-347">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="458a4-348">函数</span><span class="sxs-lookup"><span data-stu-id="458a4-348">function</span></span>||<span data-ttu-id="458a4-p123">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="458a4-p123">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="458a4-351">对象</span><span class="sxs-lookup"><span data-stu-id="458a4-351">Object</span></span>| <span data-ttu-id="458a4-352">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="458a4-352">&lt;optional&gt;</span></span>|<span data-ttu-id="458a4-353">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="458a4-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-354">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-354">Requirements</span></span>

|<span data-ttu-id="458a4-355">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-355">Requirement</span></span>| <span data-ttu-id="458a4-356">值</span><span class="sxs-lookup"><span data-stu-id="458a4-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-357">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-358">1.3</span><span class="sxs-lookup"><span data-stu-id="458a4-358">1.3</span></span>|
|[<span data-ttu-id="458a4-359">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="458a4-360">ReadItem</span></span>|
|[<span data-ttu-id="458a4-361">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-362">撰写和阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="458a4-363">示例</span><span class="sxs-lookup"><span data-stu-id="458a4-363">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="458a4-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="458a4-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="458a4-365">获取用于标识用户和 Office 外接程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="458a4-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="458a4-366">`getUserIdentityTokenAsync` 方法返回你可以用于在第三方系统上识别和 [验证外接程序和用户的令牌](https://docs.microsoft.com/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="458a4-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-367">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-367">Parameters:</span></span>

|<span data-ttu-id="458a4-368">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-368">Name</span></span>| <span data-ttu-id="458a4-369">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-369">Type</span></span>| <span data-ttu-id="458a4-370">属性</span><span class="sxs-lookup"><span data-stu-id="458a4-370">Attributes</span></span>| <span data-ttu-id="458a4-371">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="458a4-372">函数</span><span class="sxs-lookup"><span data-stu-id="458a4-372">function</span></span>||<span data-ttu-id="458a4-373">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="458a4-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="458a4-374">令牌作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="458a4-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="458a4-375">对象</span><span class="sxs-lookup"><span data-stu-id="458a4-375">Object</span></span>| <span data-ttu-id="458a4-376">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="458a4-376">&lt;optional&gt;</span></span>|<span data-ttu-id="458a4-377">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="458a4-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-378">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-378">Requirements</span></span>

|<span data-ttu-id="458a4-379">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-379">Requirement</span></span>| <span data-ttu-id="458a4-380">值</span><span class="sxs-lookup"><span data-stu-id="458a4-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-381">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-382">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-382">1.0</span></span>|
|[<span data-ttu-id="458a4-383">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-383">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="458a4-384">ReadItem</span></span>|
|[<span data-ttu-id="458a4-385">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-385">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-386">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-386">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="458a4-387">示例</span><span class="sxs-lookup"><span data-stu-id="458a4-387">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="458a4-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="458a4-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="458a4-389">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="458a4-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="458a4-390">此方法在下列应用场景不受支持。</span><span class="sxs-lookup"><span data-stu-id="458a4-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="458a4-391">在 Outlook for iOS 或 Outlook for Android 中</span><span class="sxs-lookup"><span data-stu-id="458a4-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="458a4-392">当加载项载入 Gmail 邮箱中时</span><span class="sxs-lookup"><span data-stu-id="458a4-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="458a4-393">在这些情况下，加载项应该[使用 REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) 来改为访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="458a4-393">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="458a4-394">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="458a4-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="458a4-395">有关支持的 EWS 操作的列表，请参阅[从 Outlook 加载项调用 Web 服务](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)。</span><span class="sxs-lookup"><span data-stu-id="458a4-395">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="458a4-396">你不能使用 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="458a4-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="458a4-397">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="458a4-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="458a4-p125">您的外接程序必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定访问用户邮箱的邮件外接程序的权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="458a4-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="458a4-400">服务器管理员必须在客户端访问服务器 EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="458a4-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="458a4-401">版本差异</span><span class="sxs-lookup"><span data-stu-id="458a4-401">Version differences</span></span>

<span data-ttu-id="458a4-402">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="458a4-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="458a4-p126">当邮件应用程序运行在 Outlook 网页版中时，您不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定您的邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的是 Outlook 的哪个版本。</span><span class="sxs-lookup"><span data-stu-id="458a4-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="458a4-406">参数：</span><span class="sxs-lookup"><span data-stu-id="458a4-406">Parameters:</span></span>

|<span data-ttu-id="458a4-407">名称</span><span class="sxs-lookup"><span data-stu-id="458a4-407">Name</span></span>| <span data-ttu-id="458a4-408">类型</span><span class="sxs-lookup"><span data-stu-id="458a4-408">Type</span></span>| <span data-ttu-id="458a4-409">属性</span><span class="sxs-lookup"><span data-stu-id="458a4-409">Attributes</span></span>| <span data-ttu-id="458a4-410">说明</span><span class="sxs-lookup"><span data-stu-id="458a4-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="458a4-411">字符串</span><span class="sxs-lookup"><span data-stu-id="458a4-411">String</span></span>||<span data-ttu-id="458a4-412">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="458a4-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="458a4-413">函数</span><span class="sxs-lookup"><span data-stu-id="458a4-413">function</span></span>||<span data-ttu-id="458a4-414">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="458a4-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="458a4-415">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="458a4-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="458a4-416">如果结果大小超过 1 MB，则改为返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="458a4-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="458a4-417">对象</span><span class="sxs-lookup"><span data-stu-id="458a4-417">Object</span></span>| <span data-ttu-id="458a4-418">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="458a4-418">&lt;optional&gt;</span></span>|<span data-ttu-id="458a4-419">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="458a4-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458a4-420">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-420">Requirements</span></span>

|<span data-ttu-id="458a4-421">要求</span><span class="sxs-lookup"><span data-stu-id="458a4-421">Requirement</span></span>| <span data-ttu-id="458a4-422">值</span><span class="sxs-lookup"><span data-stu-id="458a4-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="458a4-423">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="458a4-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458a4-424">1.0</span><span class="sxs-lookup"><span data-stu-id="458a4-424">1.0</span></span>|
|[<span data-ttu-id="458a4-425">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="458a4-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="458a4-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="458a4-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="458a4-427">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458a4-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458a4-428">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458a4-428">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="458a4-429">示例</span><span class="sxs-lookup"><span data-stu-id="458a4-429">Example</span></span>

<span data-ttu-id="458a4-430">下面的示例调用 `makeEwsRequestAsync` 以使用 `GetItem` 操作来获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="458a4-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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