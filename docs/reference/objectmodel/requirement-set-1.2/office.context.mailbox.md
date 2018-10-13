
# <a name="mailbox"></a><span data-ttu-id="53304-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="53304-101">mailbox</span></span>

### <span data-ttu-id="53304-p101">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="53304-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="53304-104">为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="53304-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="53304-105">要求</span><span class="sxs-lookup"><span data-stu-id="53304-105">Requirements</span></span>

|<span data-ttu-id="53304-106">要求</span><span class="sxs-lookup"><span data-stu-id="53304-106">Requirement</span></span>| <span data-ttu-id="53304-107">值</span><span class="sxs-lookup"><span data-stu-id="53304-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-108">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-109">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-109">1.0</span></span>|
|[<span data-ttu-id="53304-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-111">受限</span><span class="sxs-lookup"><span data-stu-id="53304-111">Restricted</span></span>|
|[<span data-ttu-id="53304-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-113">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="53304-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="53304-114">命名空间</span><span class="sxs-lookup"><span data-stu-id="53304-114">Namespaces</span></span>

<span data-ttu-id="53304-115">[diagnostics](Office.context.mailbox.diagnostics.md)：将诊断信息提供给 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="53304-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="53304-116">[item](Office.context.mailbox.item.md)：提供用于访问 Outlook 加载项中的邮件或约会的方法和属性。</span><span class="sxs-lookup"><span data-stu-id="53304-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="53304-117">[userProfile](Office.context.mailbox.userProfile.md)：提供有关 Outlook 加载项中的用户信息。</span><span class="sxs-lookup"><span data-stu-id="53304-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="53304-118">成员</span><span class="sxs-lookup"><span data-stu-id="53304-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="53304-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="53304-119">ewsUrl :String</span></span>

<span data-ttu-id="53304-p102">获取此电子邮件帐户的 Exchange Web 服务 (EWS) 端点 URL。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="53304-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="53304-122">注意：在 iOS 版 Outlook 或  Android 版 Outlook 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="53304-122">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="53304-p103">远程服务可使用 `ewsUrl` 值对用户邮箱进行 EWS 调用。例如，可以创建远程服务来 [获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="53304-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="53304-125">类型：</span><span class="sxs-lookup"><span data-stu-id="53304-125">Type:</span></span>

*   <span data-ttu-id="53304-126">String</span><span class="sxs-lookup"><span data-stu-id="53304-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="53304-127">要求</span><span class="sxs-lookup"><span data-stu-id="53304-127">Requirements</span></span>

|<span data-ttu-id="53304-128">要求</span><span class="sxs-lookup"><span data-stu-id="53304-128">Requirement</span></span>| <span data-ttu-id="53304-129">值</span><span class="sxs-lookup"><span data-stu-id="53304-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-130">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-131">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-131">1.0</span></span>|
|[<span data-ttu-id="53304-132">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53304-133">ReadItem</span></span>|
|[<span data-ttu-id="53304-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-135">阅读</span><span class="sxs-lookup"><span data-stu-id="53304-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="53304-136">方法</span><span class="sxs-lookup"><span data-stu-id="53304-136">Methods</span></span>

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime"></a><span data-ttu-id="53304-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="53304-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)}</span></span>

<span data-ttu-id="53304-138">获取包含以本地客户端时间表示时间信息的字典。</span><span class="sxs-lookup"><span data-stu-id="53304-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="53304-p104">Outlook 或 Outlook Web App 邮件应用程序的日期和时间可以使用不同的时区。Outlook 使用客户端计算机时区；Outlook Web App 使用 Exchange 管理中心 (EAC) 中设置的时区。应对日期和时间值进行处理，以便用户界面上显示的值始终与用户预期的时区一致。</span><span class="sxs-lookup"><span data-stu-id="53304-p104">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="53304-p105">如果在 Outlook 中运行邮件应用程序，`convertToLocalClientTime` 方法将返回多个值设置为客户端计算机时区的字典对象。如果在 Outlook Web App 中运行邮件应用程序，`convertToLocalClientTime` 方法将返回多个值设置为 EAC 中指定的时区的字典对象。</span><span class="sxs-lookup"><span data-stu-id="53304-p105">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="53304-144">参数：</span><span class="sxs-lookup"><span data-stu-id="53304-144">Parameters:</span></span>

|<span data-ttu-id="53304-145">名称</span><span class="sxs-lookup"><span data-stu-id="53304-145">Name</span></span>| <span data-ttu-id="53304-146">类型</span><span class="sxs-lookup"><span data-stu-id="53304-146">Type</span></span>| <span data-ttu-id="53304-147">说明</span><span class="sxs-lookup"><span data-stu-id="53304-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="53304-148">Date</span><span class="sxs-lookup"><span data-stu-id="53304-148">Date</span></span>|<span data-ttu-id="53304-149">一个 Date 对象</span><span class="sxs-lookup"><span data-stu-id="53304-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53304-150">要求</span><span class="sxs-lookup"><span data-stu-id="53304-150">Requirements</span></span>

|<span data-ttu-id="53304-151">要求</span><span class="sxs-lookup"><span data-stu-id="53304-151">Requirement</span></span>| <span data-ttu-id="53304-152">值</span><span class="sxs-lookup"><span data-stu-id="53304-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-153">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-154">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-154">1.0</span></span>|
|[<span data-ttu-id="53304-155">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53304-156">ReadItem</span></span>|
|[<span data-ttu-id="53304-157">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-158">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="53304-158">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="53304-159">返回：</span><span class="sxs-lookup"><span data-stu-id="53304-159">Returns:</span></span>

<span data-ttu-id="53304-160">类型：[LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="53304-160">Type: [LocalClientTime](/javascript/api/outlook_1_2/office.LocalClientTime)</span></span>

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="53304-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="53304-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="53304-162">从包含时间信息的字典中获取 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="53304-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="53304-163">`convertToUtcClientTime` 方法将包含本地日期和时间的字典转换为包含与本地日期和时间对应的正确值的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="53304-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="53304-164">参数：</span><span class="sxs-lookup"><span data-stu-id="53304-164">Parameters:</span></span>

|<span data-ttu-id="53304-165">名称</span><span class="sxs-lookup"><span data-stu-id="53304-165">Name</span></span>| <span data-ttu-id="53304-166">类型</span><span class="sxs-lookup"><span data-stu-id="53304-166">Type</span></span>| <span data-ttu-id="53304-167">说明</span><span class="sxs-lookup"><span data-stu-id="53304-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="53304-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="53304-168">LocalClientTime</span></span>](/javascript/api/outlook_1_2/office.LocalClientTime)|<span data-ttu-id="53304-169">要转换的本地时间值。</span><span class="sxs-lookup"><span data-stu-id="53304-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53304-170">要求</span><span class="sxs-lookup"><span data-stu-id="53304-170">Requirements</span></span>

|<span data-ttu-id="53304-171">要求</span><span class="sxs-lookup"><span data-stu-id="53304-171">Requirement</span></span>| <span data-ttu-id="53304-172">值</span><span class="sxs-lookup"><span data-stu-id="53304-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-173">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-173">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-174">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-174">1.0</span></span>|
|[<span data-ttu-id="53304-175">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53304-176">ReadItem</span></span>|
|[<span data-ttu-id="53304-177">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-178">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="53304-178">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="53304-179">返回：</span><span class="sxs-lookup"><span data-stu-id="53304-179">Returns:</span></span>

<span data-ttu-id="53304-180">以 UTC 格式表示时间的 Date 对象。</span><span class="sxs-lookup"><span data-stu-id="53304-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="53304-181">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="53304-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="53304-182">Date</span><span class="sxs-lookup"><span data-stu-id="53304-182">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="53304-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="53304-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="53304-184">显示现有日历约会。</span><span class="sxs-lookup"><span data-stu-id="53304-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="53304-185">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="53304-185">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="53304-186">`displayAppointmentForm` 方法将在桌面新窗口中或移动设备对话框中打开现有的日历约会。</span><span class="sxs-lookup"><span data-stu-id="53304-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="53304-p106">在 Outlook for Mac 中，您可以使用此方法来显示非重复性的单个约会，或显示重复系列中的主约会，但无法显示该系列的实例。这是因为在 Outlook for Mac 中，无法访问定期系列实例的属性（包括项目 ID）。</span><span class="sxs-lookup"><span data-stu-id="53304-p106">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="53304-189">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="53304-189">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="53304-190">如果指定的项目标识符没有识别现有约会，将在客户端计算机或设备上打开一个空白窗格，并且不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="53304-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="53304-191">参数：</span><span class="sxs-lookup"><span data-stu-id="53304-191">Parameters:</span></span>

|<span data-ttu-id="53304-192">名称</span><span class="sxs-lookup"><span data-stu-id="53304-192">Name</span></span>| <span data-ttu-id="53304-193">类型</span><span class="sxs-lookup"><span data-stu-id="53304-193">Type</span></span>| <span data-ttu-id="53304-194">说明</span><span class="sxs-lookup"><span data-stu-id="53304-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="53304-195">String</span><span class="sxs-lookup"><span data-stu-id="53304-195">String</span></span>|<span data-ttu-id="53304-196">现有日历约会的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="53304-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53304-197">要求</span><span class="sxs-lookup"><span data-stu-id="53304-197">Requirements</span></span>

|<span data-ttu-id="53304-198">要求</span><span class="sxs-lookup"><span data-stu-id="53304-198">Requirement</span></span>| <span data-ttu-id="53304-199">值</span><span class="sxs-lookup"><span data-stu-id="53304-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-200">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-200">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-201">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-201">1.0</span></span>|
|[<span data-ttu-id="53304-202">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-202">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53304-203">ReadItem</span></span>|
|[<span data-ttu-id="53304-204">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-204">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-205">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="53304-205">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="53304-206">示例</span><span class="sxs-lookup"><span data-stu-id="53304-206">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="53304-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="53304-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="53304-208">显示一封现有邮件。</span><span class="sxs-lookup"><span data-stu-id="53304-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="53304-209">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="53304-209">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="53304-210">在 Outlook Web App 中，`displayMessageForm` 方法将在桌面新窗口中或移动设备对话框中打开一封现有邮件。</span><span class="sxs-lookup"><span data-stu-id="53304-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="53304-211">在 Outlook Web App 中，此方法仅在窗体正文小于或等于 32 KB 字符数时，才会打开指定的窗体。</span><span class="sxs-lookup"><span data-stu-id="53304-211">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="53304-212">如果指定的项标识符未识别现有消息，则客户端计算机上不会显示任何消息，并且也不会返回错误消息。</span><span class="sxs-lookup"><span data-stu-id="53304-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="53304-p107">请勿将 `displayMessageForm` 与表示约会的 `itemId` 配合使用。 使用 `displayAppointmentForm` 方法显示一个现有约会， 且 `displayNewAppointmentForm` 显示一个创建新约会的窗体。</span><span class="sxs-lookup"><span data-stu-id="53304-p107">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="53304-215">参数：</span><span class="sxs-lookup"><span data-stu-id="53304-215">Parameters:</span></span>

|<span data-ttu-id="53304-216">名称</span><span class="sxs-lookup"><span data-stu-id="53304-216">Name</span></span>| <span data-ttu-id="53304-217">类型</span><span class="sxs-lookup"><span data-stu-id="53304-217">Type</span></span>| <span data-ttu-id="53304-218">说明</span><span class="sxs-lookup"><span data-stu-id="53304-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="53304-219">String</span><span class="sxs-lookup"><span data-stu-id="53304-219">String</span></span>|<span data-ttu-id="53304-220">现有邮件的 Exchange Web 服务 (EWS) 标识符。</span><span class="sxs-lookup"><span data-stu-id="53304-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53304-221">要求</span><span class="sxs-lookup"><span data-stu-id="53304-221">Requirements</span></span>

|<span data-ttu-id="53304-222">要求</span><span class="sxs-lookup"><span data-stu-id="53304-222">Requirement</span></span>| <span data-ttu-id="53304-223">值</span><span class="sxs-lookup"><span data-stu-id="53304-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-224">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-225">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-225">1.0</span></span>|
|[<span data-ttu-id="53304-226">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53304-227">ReadItem</span></span>|
|[<span data-ttu-id="53304-228">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-229">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="53304-229">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="53304-230">示例</span><span class="sxs-lookup"><span data-stu-id="53304-230">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="53304-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="53304-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="53304-232">显示用于新建日历约会的窗体。</span><span class="sxs-lookup"><span data-stu-id="53304-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="53304-233">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="53304-233">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="53304-p108">`displayNewAppointmentForm` 方法打开可让用户新建约会或会议的窗体。如果指定了参数，将使用参数的内容自动填充参数内容。</span><span class="sxs-lookup"><span data-stu-id="53304-p108">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="53304-p109">在 Outlook Web App 和适用于设备的 OWA 中，此方法始终显示包含与会者字段的窗体。如果你未将任何与会者指定为输入参数，该方法将显示为一个包含**保存**按钮的窗体。如果已指定与会者，窗体将包含与会者和“**发送**”按钮。</span><span class="sxs-lookup"><span data-stu-id="53304-p109">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="53304-p110">在 Outlook 富客户端和 Outlook RT 中，如果在 `requiredAttendees`、`optionalAttendees` 或 `resources` 参数中指定任何与会者或资源，此方法将显示会议窗体，其中包含一个“**发送**”按钮。如果未指定任何收件人，此方法将显示一个包含“**保存并关闭**”按钮的约会窗体。</span><span class="sxs-lookup"><span data-stu-id="53304-p110">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="53304-241">如果任何参数超过指定大小限制，或者指定了未知参数名称，则会引发异常。</span><span class="sxs-lookup"><span data-stu-id="53304-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="53304-242">参数：</span><span class="sxs-lookup"><span data-stu-id="53304-242">Parameters:</span></span>

|<span data-ttu-id="53304-243">名称</span><span class="sxs-lookup"><span data-stu-id="53304-243">Name</span></span>| <span data-ttu-id="53304-244">类型</span><span class="sxs-lookup"><span data-stu-id="53304-244">Type</span></span>| <span data-ttu-id="53304-245">说明</span><span class="sxs-lookup"><span data-stu-id="53304-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="53304-246">Object</span><span class="sxs-lookup"><span data-stu-id="53304-246">Object</span></span> | <span data-ttu-id="53304-247">描述新约会的参数字典。</span><span class="sxs-lookup"><span data-stu-id="53304-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="53304-248">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="53304-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="53304-p111">包含电子邮件地址的字符串数组或包含约会的每个必需与会者 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="53304-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="53304-251">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="53304-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="53304-p112">包含电子邮件地址的字符串数组或包含约会的每个可选与会者 `EmailAddressDetails` 对象的数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="53304-p112">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="53304-254">Date</span><span class="sxs-lookup"><span data-stu-id="53304-254">Date</span></span> | <span data-ttu-id="53304-255">指定约会开始日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="53304-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="53304-256">Date</span><span class="sxs-lookup"><span data-stu-id="53304-256">Date</span></span> | <span data-ttu-id="53304-257">指定约会的结束日期和时间的 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="53304-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="53304-258">String</span><span class="sxs-lookup"><span data-stu-id="53304-258">String</span></span> | <span data-ttu-id="53304-p113">包含约会位置的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="53304-p113">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="53304-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="53304-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="53304-p114">包含约会所需资源的字符串数组。数组限制为最多 100 个条目。</span><span class="sxs-lookup"><span data-stu-id="53304-p114">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="53304-264">String</span><span class="sxs-lookup"><span data-stu-id="53304-264">String</span></span> | <span data-ttu-id="53304-p115">包含约会主题的字符串。字符串长度限制为最多 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="53304-p115">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="53304-267">String</span><span class="sxs-lookup"><span data-stu-id="53304-267">String</span></span> | <span data-ttu-id="53304-p116">约会的正文。正文内容限制为最大 32 KB。</span><span class="sxs-lookup"><span data-stu-id="53304-p116">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="53304-270">要求</span><span class="sxs-lookup"><span data-stu-id="53304-270">Requirements</span></span>

|<span data-ttu-id="53304-271">要求</span><span class="sxs-lookup"><span data-stu-id="53304-271">Requirement</span></span>| <span data-ttu-id="53304-272">值</span><span class="sxs-lookup"><span data-stu-id="53304-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-273">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-274">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-274">1.0</span></span>|
|[<span data-ttu-id="53304-275">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53304-276">ReadItem</span></span>|
|[<span data-ttu-id="53304-277">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-278">阅读</span><span class="sxs-lookup"><span data-stu-id="53304-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="53304-279">示例</span><span class="sxs-lookup"><span data-stu-id="53304-279">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="53304-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="53304-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="53304-281">获取一个字符串，其中包含用于从 Exchange Server 获取附件或项目的令牌。</span><span class="sxs-lookup"><span data-stu-id="53304-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="53304-p117">`getCallbackTokenAsync` 方法进行异步调用，从托管用户邮箱的 Exchange Server 获取非跳转令牌。回调令牌的生存期为 5 分钟。</span><span class="sxs-lookup"><span data-stu-id="53304-p117">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="53304-p118">可以将令牌和附件标识符或项标识符传递至第三方系统。第三方系统使用令牌作为持有者身份验证令牌调用 Exchange Web 服务 (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) 或 [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation)，以返回附件或项目。例如，可以创建远程服务来[获取选定项目中的附件](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)。</span><span class="sxs-lookup"><span data-stu-id="53304-p118">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="53304-287">应用必须在其清单中指定拥有 **ReadItem** 权限，才能调用 `getCallbackTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="53304-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="53304-288">参数：</span><span class="sxs-lookup"><span data-stu-id="53304-288">Parameters:</span></span>

|<span data-ttu-id="53304-289">名称</span><span class="sxs-lookup"><span data-stu-id="53304-289">Name</span></span>| <span data-ttu-id="53304-290">类型</span><span class="sxs-lookup"><span data-stu-id="53304-290">Type</span></span>| <span data-ttu-id="53304-291">属性</span><span class="sxs-lookup"><span data-stu-id="53304-291">Attributes</span></span>| <span data-ttu-id="53304-292">说明</span><span class="sxs-lookup"><span data-stu-id="53304-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="53304-293">function</span><span class="sxs-lookup"><span data-stu-id="53304-293">function</span></span>||<span data-ttu-id="53304-294">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="53304-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="53304-295">令牌以 `asyncResult.value` 属性中的字符串形式提供。</span><span class="sxs-lookup"><span data-stu-id="53304-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="53304-296">Object</span><span class="sxs-lookup"><span data-stu-id="53304-296">Object</span></span>| <span data-ttu-id="53304-297">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="53304-297">&lt;optional&gt;</span></span>|<span data-ttu-id="53304-298">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="53304-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53304-299">要求</span><span class="sxs-lookup"><span data-stu-id="53304-299">Requirements</span></span>

|<span data-ttu-id="53304-300">要求</span><span class="sxs-lookup"><span data-stu-id="53304-300">Requirement</span></span>| <span data-ttu-id="53304-301">值</span><span class="sxs-lookup"><span data-stu-id="53304-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-302">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-302">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-303">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-303">1.0</span></span>|
|[<span data-ttu-id="53304-304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53304-305">ReadItem</span></span>|
|[<span data-ttu-id="53304-306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-307">阅读</span><span class="sxs-lookup"><span data-stu-id="53304-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="53304-308">示例</span><span class="sxs-lookup"><span data-stu-id="53304-308">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="53304-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="53304-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="53304-310">获取用于标识用户和 Office 加载项的令牌。</span><span class="sxs-lookup"><span data-stu-id="53304-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="53304-311">`getUserIdentityTokenAsync` 方法返回可以用于在第三方系统上识别和[验证加载项和用户的令牌](https://docs.microsoft.com/outlook/add-ins/authentication)。</span><span class="sxs-lookup"><span data-stu-id="53304-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="53304-312">参数：</span><span class="sxs-lookup"><span data-stu-id="53304-312">Parameters:</span></span>

|<span data-ttu-id="53304-313">名称</span><span class="sxs-lookup"><span data-stu-id="53304-313">Name</span></span>| <span data-ttu-id="53304-314">类型</span><span class="sxs-lookup"><span data-stu-id="53304-314">Type</span></span>| <span data-ttu-id="53304-315">属性</span><span class="sxs-lookup"><span data-stu-id="53304-315">Attributes</span></span>| <span data-ttu-id="53304-316">说明</span><span class="sxs-lookup"><span data-stu-id="53304-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="53304-317">function</span><span class="sxs-lookup"><span data-stu-id="53304-317">function</span></span>||<span data-ttu-id="53304-318">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="53304-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="53304-319">令牌以 `asyncResult.value` 属性中的字符串形式提供。</span><span class="sxs-lookup"><span data-stu-id="53304-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="53304-320">Object</span><span class="sxs-lookup"><span data-stu-id="53304-320">Object</span></span>| <span data-ttu-id="53304-321">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="53304-321">&lt;optional&gt;</span></span>|<span data-ttu-id="53304-322">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="53304-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53304-323">要求</span><span class="sxs-lookup"><span data-stu-id="53304-323">Requirements</span></span>

|<span data-ttu-id="53304-324">要求</span><span class="sxs-lookup"><span data-stu-id="53304-324">Requirement</span></span>| <span data-ttu-id="53304-325">值</span><span class="sxs-lookup"><span data-stu-id="53304-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-326">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-326">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-327">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-327">1.0</span></span>|
|[<span data-ttu-id="53304-328">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53304-329">ReadItem</span></span>|
|[<span data-ttu-id="53304-330">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-331">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="53304-331">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="53304-332">示例</span><span class="sxs-lookup"><span data-stu-id="53304-332">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="53304-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="53304-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="53304-334">向托管用户邮箱的 Exchange 服务器上的 Exchange Web 服务 (EWS) 发出异步请求。</span><span class="sxs-lookup"><span data-stu-id="53304-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="53304-335">在以下情景中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="53304-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="53304-336">在 Outlook for iOS 或 Outlook for Android 中</span><span class="sxs-lookup"><span data-stu-id="53304-336">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="53304-337">通过 Gmail 邮箱加载加载项</span><span class="sxs-lookup"><span data-stu-id="53304-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="53304-338">在这些情况下, 加载项应转而[使用 REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) 访问用户的邮箱。</span><span class="sxs-lookup"><span data-stu-id="53304-338">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="53304-339">`makeEwsRequestAsync` 方法代表加载项将 EWS 请求发送到 Exchange。</span><span class="sxs-lookup"><span data-stu-id="53304-339">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="53304-340">有关支持的 EWS 操作列表，请参阅[从 Outlook 加载项调用 Web 服务](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) 。</span><span class="sxs-lookup"><span data-stu-id="53304-340">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="53304-341">不能通过 `makeEwsRequestAsync` 方法请求与文件夹关联的项目。</span><span class="sxs-lookup"><span data-stu-id="53304-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="53304-342">XML 请求必须指定 UTF-8 编码。</span><span class="sxs-lookup"><span data-stu-id="53304-342">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="53304-p120">加载项必须具有 **ReadWriteMailbox** 权限才能使用 `makeEwsRequestAsync` 方法。有关使用 **ReadWriteMailbox** 权限和可使用 `makeEwsRequestAsync` 方法调用 EWS 操作的信息，请参阅[指定用户邮箱中邮件加载项的访问权限](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)。</span><span class="sxs-lookup"><span data-stu-id="53304-p120">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="53304-345">服务器管理员必须在 Client Access Server EWS 目录上将 `OAuthAuthentication` 设置为 true，`makeEwsRequestAsync` 方法才能发出 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="53304-345">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="53304-346">版本差异</span><span class="sxs-lookup"><span data-stu-id="53304-346">Version differences</span></span>

<span data-ttu-id="53304-347">当你在较 15.0.4535.1004 版本更早的 Outlook 版本中运行的邮件应用程序中使用 `makeEwsRequestAsync` 方法时，应当将编码值设置为 `ISO-8859-1`。</span><span class="sxs-lookup"><span data-stu-id="53304-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="53304-p121">当邮件应用程序在 Outlook 网页版中运行时，不需要设置编码值。可以通过使用 mailbox.diagnostics.hostName 属性来确定邮件应用程序在 Outlook 中还是 Outlook 网页版中运行。可以通过使用 mailbox.diagnostics.hostVersion 属性来确定正在运行的 Outlook 版本。</span><span class="sxs-lookup"><span data-stu-id="53304-p121">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="53304-351">参数：</span><span class="sxs-lookup"><span data-stu-id="53304-351">Parameters:</span></span>

|<span data-ttu-id="53304-352">名称</span><span class="sxs-lookup"><span data-stu-id="53304-352">Name</span></span>| <span data-ttu-id="53304-353">类型</span><span class="sxs-lookup"><span data-stu-id="53304-353">Type</span></span>| <span data-ttu-id="53304-354">属性</span><span class="sxs-lookup"><span data-stu-id="53304-354">Attributes</span></span>| <span data-ttu-id="53304-355">说明</span><span class="sxs-lookup"><span data-stu-id="53304-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="53304-356">String</span><span class="sxs-lookup"><span data-stu-id="53304-356">String</span></span>||<span data-ttu-id="53304-357">EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="53304-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="53304-358">function</span><span class="sxs-lookup"><span data-stu-id="53304-358">function</span></span>||<span data-ttu-id="53304-359">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="53304-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="53304-360">EWS 调用的 XML 结果作为 `asyncResult.value` 属性中的字符串提供。</span><span class="sxs-lookup"><span data-stu-id="53304-360">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="53304-361">如果结果的大小超过 1 MB，将转而返回一条错误消息。</span><span class="sxs-lookup"><span data-stu-id="53304-361">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="53304-362">Object</span><span class="sxs-lookup"><span data-stu-id="53304-362">Object</span></span>| <span data-ttu-id="53304-363">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="53304-363">&lt;optional&gt;</span></span>|<span data-ttu-id="53304-364">传递给异步方法的任何状态数据。</span><span class="sxs-lookup"><span data-stu-id="53304-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53304-365">要求</span><span class="sxs-lookup"><span data-stu-id="53304-365">Requirements</span></span>

|<span data-ttu-id="53304-366">要求</span><span class="sxs-lookup"><span data-stu-id="53304-366">Requirement</span></span>| <span data-ttu-id="53304-367">值</span><span class="sxs-lookup"><span data-stu-id="53304-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="53304-368">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="53304-368">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53304-369">1.0</span><span class="sxs-lookup"><span data-stu-id="53304-369">1.0</span></span>|
|[<span data-ttu-id="53304-370">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="53304-370">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53304-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="53304-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="53304-372">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="53304-372">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="53304-373">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="53304-373">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="53304-374">示例</span><span class="sxs-lookup"><span data-stu-id="53304-374">Example</span></span>

<span data-ttu-id="53304-375">下面的示例调用 `makeEwsRequestAsync`  以使用  `GetItem` 操作获取项目的主题。</span><span class="sxs-lookup"><span data-stu-id="53304-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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