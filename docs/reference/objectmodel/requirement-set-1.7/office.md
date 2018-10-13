 

# <a name="office"></a><span data-ttu-id="35a70-101">Office</span><span class="sxs-lookup"><span data-stu-id="35a70-101">Office</span></span>

<span data-ttu-id="35a70-p101">Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的那些接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="35a70-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="35a70-104">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-104">Requirements</span></span>

|<span data-ttu-id="35a70-105">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-105">Requirement</span></span>| <span data-ttu-id="35a70-106">值</span><span class="sxs-lookup"><span data-stu-id="35a70-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a70-107">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="35a70-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a70-108">1.0</span><span class="sxs-lookup"><span data-stu-id="35a70-108">1.0</span></span>|
|[<span data-ttu-id="35a70-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35a70-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a70-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="35a70-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="35a70-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="35a70-111">Members and methods</span></span>

| <span data-ttu-id="35a70-112">成员</span><span class="sxs-lookup"><span data-stu-id="35a70-112">Member</span></span> | <span data-ttu-id="35a70-113">类型</span><span class="sxs-lookup"><span data-stu-id="35a70-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="35a70-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="35a70-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="35a70-115">成员</span><span class="sxs-lookup"><span data-stu-id="35a70-115">Member</span></span> |
| [<span data-ttu-id="35a70-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="35a70-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="35a70-117">成员</span><span class="sxs-lookup"><span data-stu-id="35a70-117">Member</span></span> |
| [<span data-ttu-id="35a70-118">EventType</span><span class="sxs-lookup"><span data-stu-id="35a70-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="35a70-119">成员</span><span class="sxs-lookup"><span data-stu-id="35a70-119">Member</span></span> |
| [<span data-ttu-id="35a70-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="35a70-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="35a70-121">成员</span><span class="sxs-lookup"><span data-stu-id="35a70-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="35a70-122">命名空间</span><span class="sxs-lookup"><span data-stu-id="35a70-122">Namespaces</span></span>

<span data-ttu-id="35a70-123">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="35a70-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="35a70-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="35a70-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="35a70-125">成员</span><span class="sxs-lookup"><span data-stu-id="35a70-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="35a70-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="35a70-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="35a70-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="35a70-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="35a70-128">类型：</span><span class="sxs-lookup"><span data-stu-id="35a70-128">Type:</span></span>

*   <span data-ttu-id="35a70-129">String</span><span class="sxs-lookup"><span data-stu-id="35a70-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35a70-130">属性：</span><span class="sxs-lookup"><span data-stu-id="35a70-130">Properties:</span></span>

|<span data-ttu-id="35a70-131">名称</span><span class="sxs-lookup"><span data-stu-id="35a70-131">Name</span></span>| <span data-ttu-id="35a70-132">类型</span><span class="sxs-lookup"><span data-stu-id="35a70-132">Type</span></span>| <span data-ttu-id="35a70-133">说明</span><span class="sxs-lookup"><span data-stu-id="35a70-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="35a70-134">String</span><span class="sxs-lookup"><span data-stu-id="35a70-134">String</span></span>|<span data-ttu-id="35a70-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="35a70-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="35a70-136">String</span><span class="sxs-lookup"><span data-stu-id="35a70-136">String</span></span>|<span data-ttu-id="35a70-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="35a70-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a70-138">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-138">Requirements</span></span>

|<span data-ttu-id="35a70-139">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-139">Requirement</span></span>| <span data-ttu-id="35a70-140">值</span><span class="sxs-lookup"><span data-stu-id="35a70-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a70-141">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="35a70-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a70-142">1.0</span><span class="sxs-lookup"><span data-stu-id="35a70-142">1.0</span></span>|
|[<span data-ttu-id="35a70-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35a70-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a70-144">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="35a70-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="35a70-145">CoercionType :字符串</span><span class="sxs-lookup"><span data-stu-id="35a70-145">CoercionType :String</span></span>

<span data-ttu-id="35a70-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="35a70-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="35a70-147">类型：</span><span class="sxs-lookup"><span data-stu-id="35a70-147">Type:</span></span>

*   <span data-ttu-id="35a70-148">String</span><span class="sxs-lookup"><span data-stu-id="35a70-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35a70-149">属性：</span><span class="sxs-lookup"><span data-stu-id="35a70-149">Properties:</span></span>

|<span data-ttu-id="35a70-150">名称</span><span class="sxs-lookup"><span data-stu-id="35a70-150">Name</span></span>| <span data-ttu-id="35a70-151">类型</span><span class="sxs-lookup"><span data-stu-id="35a70-151">Type</span></span>| <span data-ttu-id="35a70-152">说明</span><span class="sxs-lookup"><span data-stu-id="35a70-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="35a70-153">String</span><span class="sxs-lookup"><span data-stu-id="35a70-153">String</span></span>|<span data-ttu-id="35a70-154">要求以 HTML 格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="35a70-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="35a70-155">String</span><span class="sxs-lookup"><span data-stu-id="35a70-155">String</span></span>|<span data-ttu-id="35a70-156">要求以文本格式返回数据。</span><span class="sxs-lookup"><span data-stu-id="35a70-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a70-157">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-157">Requirements</span></span>

|<span data-ttu-id="35a70-158">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-158">Requirement</span></span>| <span data-ttu-id="35a70-159">值</span><span class="sxs-lookup"><span data-stu-id="35a70-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a70-160">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="35a70-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a70-161">1.0</span><span class="sxs-lookup"><span data-stu-id="35a70-161">1.0</span></span>|
|[<span data-ttu-id="35a70-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35a70-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a70-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="35a70-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="35a70-164">EventType :字符串</span><span class="sxs-lookup"><span data-stu-id="35a70-164">EventType :String</span></span>

<span data-ttu-id="35a70-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="35a70-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="35a70-166">类型：</span><span class="sxs-lookup"><span data-stu-id="35a70-166">Type:</span></span>

*   <span data-ttu-id="35a70-167">String</span><span class="sxs-lookup"><span data-stu-id="35a70-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35a70-168">属性：</span><span class="sxs-lookup"><span data-stu-id="35a70-168">Properties:</span></span>

| <span data-ttu-id="35a70-169">名称</span><span class="sxs-lookup"><span data-stu-id="35a70-169">Name</span></span> | <span data-ttu-id="35a70-170">类型</span><span class="sxs-lookup"><span data-stu-id="35a70-170">Type</span></span> | <span data-ttu-id="35a70-171">说明</span><span class="sxs-lookup"><span data-stu-id="35a70-171">Description</span></span> | <span data-ttu-id="35a70-172">最低要求集</span><span class="sxs-lookup"><span data-stu-id="35a70-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="35a70-173">String</span><span class="sxs-lookup"><span data-stu-id="35a70-173">String</span></span> | <span data-ttu-id="35a70-174">所选约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="35a70-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="35a70-175">1.7</span><span class="sxs-lookup"><span data-stu-id="35a70-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="35a70-176">String</span><span class="sxs-lookup"><span data-stu-id="35a70-176">String</span></span> | <span data-ttu-id="35a70-177">选定的项已更改。</span><span class="sxs-lookup"><span data-stu-id="35a70-177">The selected item has changed.</span></span> | <span data-ttu-id="35a70-178">1.5</span><span class="sxs-lookup"><span data-stu-id="35a70-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="35a70-179">String</span><span class="sxs-lookup"><span data-stu-id="35a70-179">String</span></span> | <span data-ttu-id="35a70-180">所选项或约会地点的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="35a70-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="35a70-181">1.7</span><span class="sxs-lookup"><span data-stu-id="35a70-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="35a70-182">String</span><span class="sxs-lookup"><span data-stu-id="35a70-182">String</span></span> | <span data-ttu-id="35a70-183">所选系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="35a70-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="35a70-184">1.7</span><span class="sxs-lookup"><span data-stu-id="35a70-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35a70-185">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-185">Requirements</span></span>

|<span data-ttu-id="35a70-186">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-186">Requirement</span></span>| <span data-ttu-id="35a70-187">值</span><span class="sxs-lookup"><span data-stu-id="35a70-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a70-188">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="35a70-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a70-189">1.5</span><span class="sxs-lookup"><span data-stu-id="35a70-189">1.5</span></span> |
|[<span data-ttu-id="35a70-190">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35a70-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a70-191">撰写或阅读​</span><span class="sxs-lookup"><span data-stu-id="35a70-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="35a70-192">SourceProperty :字符串</span><span class="sxs-lookup"><span data-stu-id="35a70-192">SourceProperty :String</span></span>

<span data-ttu-id="35a70-193">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="35a70-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="35a70-194">类型：</span><span class="sxs-lookup"><span data-stu-id="35a70-194">Type:</span></span>

*   <span data-ttu-id="35a70-195">String</span><span class="sxs-lookup"><span data-stu-id="35a70-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="35a70-196">属性：</span><span class="sxs-lookup"><span data-stu-id="35a70-196">Properties:</span></span>

|<span data-ttu-id="35a70-197">名称</span><span class="sxs-lookup"><span data-stu-id="35a70-197">Name</span></span>| <span data-ttu-id="35a70-198">类型</span><span class="sxs-lookup"><span data-stu-id="35a70-198">Type</span></span>| <span data-ttu-id="35a70-199">说明</span><span class="sxs-lookup"><span data-stu-id="35a70-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="35a70-200">String</span><span class="sxs-lookup"><span data-stu-id="35a70-200">String</span></span>|<span data-ttu-id="35a70-201">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="35a70-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="35a70-202">String</span><span class="sxs-lookup"><span data-stu-id="35a70-202">String</span></span>|<span data-ttu-id="35a70-203">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="35a70-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a70-204">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-204">Requirements</span></span>

|<span data-ttu-id="35a70-205">要求</span><span class="sxs-lookup"><span data-stu-id="35a70-205">Requirement</span></span>| <span data-ttu-id="35a70-206">值</span><span class="sxs-lookup"><span data-stu-id="35a70-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a70-207">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="35a70-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a70-208">1.0</span><span class="sxs-lookup"><span data-stu-id="35a70-208">1.0</span></span>|
|[<span data-ttu-id="35a70-209">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="35a70-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a70-210">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="35a70-210">Compose or read</span></span>|