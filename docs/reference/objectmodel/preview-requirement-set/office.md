 

# <a name="office"></a><span data-ttu-id="458fa-101">Office</span><span class="sxs-lookup"><span data-stu-id="458fa-101">Office</span></span>

<span data-ttu-id="458fa-p101">Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="458fa-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="458fa-104">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-104">Requirements</span></span>

|<span data-ttu-id="458fa-105">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-105">Requirement</span></span>| <span data-ttu-id="458fa-106">值</span><span class="sxs-lookup"><span data-stu-id="458fa-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="458fa-107">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="458fa-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458fa-108">1.0</span><span class="sxs-lookup"><span data-stu-id="458fa-108">1.0</span></span>|
|[<span data-ttu-id="458fa-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458fa-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458fa-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458fa-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="458fa-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="458fa-111">Members and methods</span></span>

| <span data-ttu-id="458fa-112">成员</span><span class="sxs-lookup"><span data-stu-id="458fa-112">Member</span></span> | <span data-ttu-id="458fa-113">类型</span><span class="sxs-lookup"><span data-stu-id="458fa-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="458fa-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="458fa-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="458fa-115">成员</span><span class="sxs-lookup"><span data-stu-id="458fa-115">Member</span></span> |
| [<span data-ttu-id="458fa-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="458fa-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="458fa-117">成员</span><span class="sxs-lookup"><span data-stu-id="458fa-117">Member</span></span> |
| [<span data-ttu-id="458fa-118">EventType</span><span class="sxs-lookup"><span data-stu-id="458fa-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="458fa-119">成员</span><span class="sxs-lookup"><span data-stu-id="458fa-119">Member</span></span> |
| [<span data-ttu-id="458fa-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="458fa-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="458fa-121">成员</span><span class="sxs-lookup"><span data-stu-id="458fa-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="458fa-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="458fa-122">Namespaces</span></span>

<span data-ttu-id="458fa-123">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="458fa-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="458fa-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="458fa-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="458fa-125">成员</span><span class="sxs-lookup"><span data-stu-id="458fa-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="458fa-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="458fa-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="458fa-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="458fa-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="458fa-128">类型：</span><span class="sxs-lookup"><span data-stu-id="458fa-128">Type:</span></span>

*   <span data-ttu-id="458fa-129">String</span><span class="sxs-lookup"><span data-stu-id="458fa-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="458fa-130">属性：</span><span class="sxs-lookup"><span data-stu-id="458fa-130">Properties:</span></span>

|<span data-ttu-id="458fa-131">名称</span><span class="sxs-lookup"><span data-stu-id="458fa-131">Name</span></span>| <span data-ttu-id="458fa-132">类型</span><span class="sxs-lookup"><span data-stu-id="458fa-132">Type</span></span>| <span data-ttu-id="458fa-133">说明</span><span class="sxs-lookup"><span data-stu-id="458fa-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="458fa-134">String</span><span class="sxs-lookup"><span data-stu-id="458fa-134">String</span></span>|<span data-ttu-id="458fa-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="458fa-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="458fa-136">String</span><span class="sxs-lookup"><span data-stu-id="458fa-136">String</span></span>|<span data-ttu-id="458fa-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="458fa-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458fa-138">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-138">Requirements</span></span>

|<span data-ttu-id="458fa-139">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-139">Requirement</span></span>| <span data-ttu-id="458fa-140">值</span><span class="sxs-lookup"><span data-stu-id="458fa-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="458fa-141">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="458fa-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458fa-142">1.0</span><span class="sxs-lookup"><span data-stu-id="458fa-142">1.0</span></span>|
|[<span data-ttu-id="458fa-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458fa-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458fa-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458fa-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="458fa-145">CoercionType :字符串</span><span class="sxs-lookup"><span data-stu-id="458fa-145">CoercionType :String</span></span>

<span data-ttu-id="458fa-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="458fa-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="458fa-147">类型：</span><span class="sxs-lookup"><span data-stu-id="458fa-147">Type:</span></span>

*   <span data-ttu-id="458fa-148">String</span><span class="sxs-lookup"><span data-stu-id="458fa-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="458fa-149">属性：</span><span class="sxs-lookup"><span data-stu-id="458fa-149">Properties:</span></span>

|<span data-ttu-id="458fa-150">名称</span><span class="sxs-lookup"><span data-stu-id="458fa-150">Name</span></span>| <span data-ttu-id="458fa-151">类型</span><span class="sxs-lookup"><span data-stu-id="458fa-151">Type</span></span>| <span data-ttu-id="458fa-152">说明</span><span class="sxs-lookup"><span data-stu-id="458fa-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="458fa-153">String</span><span class="sxs-lookup"><span data-stu-id="458fa-153">String</span></span>|<span data-ttu-id="458fa-154">要求以 HTML 格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="458fa-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="458fa-155">String</span><span class="sxs-lookup"><span data-stu-id="458fa-155">String</span></span>|<span data-ttu-id="458fa-156">要求以文本格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="458fa-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458fa-157">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-157">Requirements</span></span>

|<span data-ttu-id="458fa-158">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-158">Requirement</span></span>| <span data-ttu-id="458fa-159">值</span><span class="sxs-lookup"><span data-stu-id="458fa-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="458fa-160">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="458fa-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458fa-161">1.0</span><span class="sxs-lookup"><span data-stu-id="458fa-161">1.0</span></span>|
|[<span data-ttu-id="458fa-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458fa-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458fa-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458fa-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="458fa-164">EventType :字符串</span><span class="sxs-lookup"><span data-stu-id="458fa-164">EventType :String</span></span>

<span data-ttu-id="458fa-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="458fa-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="458fa-166">类型：</span><span class="sxs-lookup"><span data-stu-id="458fa-166">Type:</span></span>

*   <span data-ttu-id="458fa-167">String</span><span class="sxs-lookup"><span data-stu-id="458fa-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="458fa-168">属性：</span><span class="sxs-lookup"><span data-stu-id="458fa-168">Properties:</span></span>

| <span data-ttu-id="458fa-169">名称</span><span class="sxs-lookup"><span data-stu-id="458fa-169">Name</span></span> | <span data-ttu-id="458fa-170">类型</span><span class="sxs-lookup"><span data-stu-id="458fa-170">Type</span></span> | <span data-ttu-id="458fa-171">说明</span><span class="sxs-lookup"><span data-stu-id="458fa-171">Description</span></span> | <span data-ttu-id="458fa-172">最低要求集</span><span class="sxs-lookup"><span data-stu-id="458fa-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="458fa-173">String</span><span class="sxs-lookup"><span data-stu-id="458fa-173">String</span></span> | <span data-ttu-id="458fa-174">所选约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="458fa-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="458fa-175">1.7</span><span class="sxs-lookup"><span data-stu-id="458fa-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="458fa-176">String</span><span class="sxs-lookup"><span data-stu-id="458fa-176">String</span></span> | <span data-ttu-id="458fa-177">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="458fa-177">The selected item has changed.</span></span> | <span data-ttu-id="458fa-178">1.5</span><span class="sxs-lookup"><span data-stu-id="458fa-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="458fa-179">String</span><span class="sxs-lookup"><span data-stu-id="458fa-179">String</span></span> | <span data-ttu-id="458fa-180">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="458fa-180">The selected item has changed.</span></span> | <span data-ttu-id="458fa-181">预览</span><span class="sxs-lookup"><span data-stu-id="458fa-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="458fa-182">String</span><span class="sxs-lookup"><span data-stu-id="458fa-182">String</span></span> | <span data-ttu-id="458fa-183">所选项或约会地点的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="458fa-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="458fa-184">1.7</span><span class="sxs-lookup"><span data-stu-id="458fa-184">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="458fa-185">String</span><span class="sxs-lookup"><span data-stu-id="458fa-185">String</span></span> | <span data-ttu-id="458fa-186">所选系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="458fa-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="458fa-187">1.7</span><span class="sxs-lookup"><span data-stu-id="458fa-187">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="458fa-188">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-188">Requirements</span></span>

|<span data-ttu-id="458fa-189">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-189">Requirement</span></span>| <span data-ttu-id="458fa-190">值</span><span class="sxs-lookup"><span data-stu-id="458fa-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="458fa-191">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="458fa-191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458fa-192">1.5</span><span class="sxs-lookup"><span data-stu-id="458fa-192">1.5</span></span> |
|[<span data-ttu-id="458fa-193">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458fa-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458fa-194">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458fa-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="458fa-195">SourceProperty :字符串</span><span class="sxs-lookup"><span data-stu-id="458fa-195">SourceProperty :String</span></span>

<span data-ttu-id="458fa-196">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="458fa-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="458fa-197">类型：</span><span class="sxs-lookup"><span data-stu-id="458fa-197">Type:</span></span>

*   <span data-ttu-id="458fa-198">String</span><span class="sxs-lookup"><span data-stu-id="458fa-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="458fa-199">属性：</span><span class="sxs-lookup"><span data-stu-id="458fa-199">Properties:</span></span>

|<span data-ttu-id="458fa-200">名称</span><span class="sxs-lookup"><span data-stu-id="458fa-200">Name</span></span>| <span data-ttu-id="458fa-201">类型</span><span class="sxs-lookup"><span data-stu-id="458fa-201">Type</span></span>| <span data-ttu-id="458fa-202">说明</span><span class="sxs-lookup"><span data-stu-id="458fa-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="458fa-203">String</span><span class="sxs-lookup"><span data-stu-id="458fa-203">String</span></span>|<span data-ttu-id="458fa-204">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="458fa-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="458fa-205">String</span><span class="sxs-lookup"><span data-stu-id="458fa-205">String</span></span>|<span data-ttu-id="458fa-206">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="458fa-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="458fa-207">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-207">Requirements</span></span>

|<span data-ttu-id="458fa-208">要求</span><span class="sxs-lookup"><span data-stu-id="458fa-208">Requirement</span></span>| <span data-ttu-id="458fa-209">值</span><span class="sxs-lookup"><span data-stu-id="458fa-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="458fa-210">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="458fa-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="458fa-211">1.0</span><span class="sxs-lookup"><span data-stu-id="458fa-211">1.0</span></span>|
|[<span data-ttu-id="458fa-212">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="458fa-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="458fa-213">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="458fa-213">Compose or read</span></span>|