 

# <a name="office"></a><span data-ttu-id="68b26-101">Office</span><span class="sxs-lookup"><span data-stu-id="68b26-101">Office</span></span>

<span data-ttu-id="68b26-p101">该 Office 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office 命名空间的完整列表，请参阅[共享 API](/javascript/api/office)。</span><span class="sxs-lookup"><span data-stu-id="68b26-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="68b26-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="68b26-104">Requirements</span></span>

|<span data-ttu-id="68b26-105">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-105">Requirement</span></span>| <span data-ttu-id="68b26-106">值</span><span class="sxs-lookup"><span data-stu-id="68b26-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="68b26-107">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="68b26-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68b26-108">1.0</span><span class="sxs-lookup"><span data-stu-id="68b26-108">1.0</span></span>|
|[<span data-ttu-id="68b26-109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="68b26-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="68b26-110">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="68b26-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="68b26-111">成员和方法</span><span class="sxs-lookup"><span data-stu-id="68b26-111">Members and methods</span></span>

| <span data-ttu-id="68b26-112">成员</span><span class="sxs-lookup"><span data-stu-id="68b26-112">Member</span></span> | <span data-ttu-id="68b26-113">类型</span><span class="sxs-lookup"><span data-stu-id="68b26-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="68b26-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="68b26-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="68b26-115">成员</span><span class="sxs-lookup"><span data-stu-id="68b26-115">Member</span></span> |
| [<span data-ttu-id="68b26-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="68b26-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="68b26-117">成员</span><span class="sxs-lookup"><span data-stu-id="68b26-117">Member</span></span> |
| [<span data-ttu-id="68b26-118">EventType</span><span class="sxs-lookup"><span data-stu-id="68b26-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="68b26-119">成员</span><span class="sxs-lookup"><span data-stu-id="68b26-119">Member</span></span> |
| [<span data-ttu-id="68b26-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="68b26-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="68b26-121">成员</span><span class="sxs-lookup"><span data-stu-id="68b26-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="68b26-122">命名空间</span><span class="sxs-lookup"><span data-stu-id="68b26-122">Namespaces</span></span>

<span data-ttu-id="68b26-123">[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。</span><span class="sxs-lookup"><span data-stu-id="68b26-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="68b26-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。</span><span class="sxs-lookup"><span data-stu-id="68b26-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="68b26-125">成员</span><span class="sxs-lookup"><span data-stu-id="68b26-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="68b26-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="68b26-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="68b26-127">指定异步调用的结果。</span><span class="sxs-lookup"><span data-stu-id="68b26-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="68b26-128">类型：</span><span class="sxs-lookup"><span data-stu-id="68b26-128">Type:</span></span>

*   <span data-ttu-id="68b26-129">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="68b26-130">属性：</span><span class="sxs-lookup"><span data-stu-id="68b26-130">Properties:</span></span>

|<span data-ttu-id="68b26-131">名称</span><span class="sxs-lookup"><span data-stu-id="68b26-131">Name</span></span>| <span data-ttu-id="68b26-132">类型</span><span class="sxs-lookup"><span data-stu-id="68b26-132">Type</span></span>| <span data-ttu-id="68b26-133">描述</span><span class="sxs-lookup"><span data-stu-id="68b26-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="68b26-134">String</span><span class="sxs-lookup"><span data-stu-id="68b26-134">String</span></span>|<span data-ttu-id="68b26-135">调用成功。</span><span class="sxs-lookup"><span data-stu-id="68b26-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="68b26-136">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-136">String</span></span>|<span data-ttu-id="68b26-137">调用失败。</span><span class="sxs-lookup"><span data-stu-id="68b26-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68b26-138">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-138">Requirements</span></span>

|<span data-ttu-id="68b26-139">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-139">Requirement</span></span>| <span data-ttu-id="68b26-140">值</span><span class="sxs-lookup"><span data-stu-id="68b26-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="68b26-141">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="68b26-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68b26-142">1.0</span><span class="sxs-lookup"><span data-stu-id="68b26-142">1.0</span></span>|
|[<span data-ttu-id="68b26-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="68b26-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="68b26-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="68b26-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="68b26-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="68b26-145">CoercionType :String</span></span>

<span data-ttu-id="68b26-146">指定如何强制由调用方法返回或设置的数据。</span><span class="sxs-lookup"><span data-stu-id="68b26-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="68b26-147">类型：</span><span class="sxs-lookup"><span data-stu-id="68b26-147">Type:</span></span>

*   <span data-ttu-id="68b26-148">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="68b26-149">属性：</span><span class="sxs-lookup"><span data-stu-id="68b26-149">Properties:</span></span>

|<span data-ttu-id="68b26-150">名称</span><span class="sxs-lookup"><span data-stu-id="68b26-150">Name</span></span>| <span data-ttu-id="68b26-151">类型</span><span class="sxs-lookup"><span data-stu-id="68b26-151">Type</span></span>| <span data-ttu-id="68b26-152">描述</span><span class="sxs-lookup"><span data-stu-id="68b26-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="68b26-153">String</span><span class="sxs-lookup"><span data-stu-id="68b26-153">String</span></span>|<span data-ttu-id="68b26-154">请求以 HTML 格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="68b26-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="68b26-155">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-155">String</span></span>|<span data-ttu-id="68b26-156">请求以文本格式返回的数据。</span><span class="sxs-lookup"><span data-stu-id="68b26-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68b26-157">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-157">Requirements</span></span>

|<span data-ttu-id="68b26-158">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-158">Requirement</span></span>| <span data-ttu-id="68b26-159">值</span><span class="sxs-lookup"><span data-stu-id="68b26-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="68b26-160">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="68b26-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68b26-161">1.0</span><span class="sxs-lookup"><span data-stu-id="68b26-161">1.0</span></span>|
|[<span data-ttu-id="68b26-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="68b26-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="68b26-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="68b26-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="68b26-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="68b26-164">EventType :String</span></span>

<span data-ttu-id="68b26-165">指定与事件处理程序相关联的事件。</span><span class="sxs-lookup"><span data-stu-id="68b26-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="68b26-166">类型：</span><span class="sxs-lookup"><span data-stu-id="68b26-166">Type:</span></span>

*   <span data-ttu-id="68b26-167">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="68b26-168">属性：</span><span class="sxs-lookup"><span data-stu-id="68b26-168">Properties:</span></span>

| <span data-ttu-id="68b26-169">名称</span><span class="sxs-lookup"><span data-stu-id="68b26-169">Name</span></span> | <span data-ttu-id="68b26-170">类型</span><span class="sxs-lookup"><span data-stu-id="68b26-170">Type</span></span> | <span data-ttu-id="68b26-171">描述</span><span class="sxs-lookup"><span data-stu-id="68b26-171">Description</span></span> | <span data-ttu-id="68b26-172">最低要求集</span><span class="sxs-lookup"><span data-stu-id="68b26-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="68b26-173">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-173">String</span></span> | <span data-ttu-id="68b26-174">所选的约会或系列的日期或时间已更改。</span><span class="sxs-lookup"><span data-stu-id="68b26-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="68b26-175">1.7</span><span class="sxs-lookup"><span data-stu-id="68b26-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="68b26-176">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-176">String</span></span> | <span data-ttu-id="68b26-177">在任务窗格固定时，将选择不同的 Outlook 项进行查看。</span><span class="sxs-lookup"><span data-stu-id="68b26-177">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="68b26-178">1.5</span><span class="sxs-lookup"><span data-stu-id="68b26-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="68b26-179">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-179">String</span></span> | <span data-ttu-id="68b26-180">选定项目或约会位置的收件人列表已更改。</span><span class="sxs-lookup"><span data-stu-id="68b26-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="68b26-181">1.7</span><span class="sxs-lookup"><span data-stu-id="68b26-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="68b26-182">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-182">String</span></span> | <span data-ttu-id="68b26-183">选定系列的定期模式已更改。</span><span class="sxs-lookup"><span data-stu-id="68b26-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="68b26-184">1.7</span><span class="sxs-lookup"><span data-stu-id="68b26-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="68b26-185">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-185">Requirements</span></span>

|<span data-ttu-id="68b26-186">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-186">Requirement</span></span>| <span data-ttu-id="68b26-187">值</span><span class="sxs-lookup"><span data-stu-id="68b26-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="68b26-188">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="68b26-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68b26-189">1.5</span><span class="sxs-lookup"><span data-stu-id="68b26-189">1.5</span></span> |
|[<span data-ttu-id="68b26-190">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="68b26-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="68b26-191">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="68b26-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="68b26-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="68b26-192">SourceProperty :String</span></span>

<span data-ttu-id="68b26-193">指定由调用方法返回的数据源。</span><span class="sxs-lookup"><span data-stu-id="68b26-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="68b26-194">类型：</span><span class="sxs-lookup"><span data-stu-id="68b26-194">Type:</span></span>

*   <span data-ttu-id="68b26-195">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="68b26-196">属性：</span><span class="sxs-lookup"><span data-stu-id="68b26-196">Properties:</span></span>

|<span data-ttu-id="68b26-197">名称</span><span class="sxs-lookup"><span data-stu-id="68b26-197">Name</span></span>| <span data-ttu-id="68b26-198">类型</span><span class="sxs-lookup"><span data-stu-id="68b26-198">Type</span></span>| <span data-ttu-id="68b26-199">描述</span><span class="sxs-lookup"><span data-stu-id="68b26-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="68b26-200">字符串</span><span class="sxs-lookup"><span data-stu-id="68b26-200">String</span></span>|<span data-ttu-id="68b26-201">数据源来自邮件的正文。</span><span class="sxs-lookup"><span data-stu-id="68b26-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="68b26-202">String</span><span class="sxs-lookup"><span data-stu-id="68b26-202">String</span></span>|<span data-ttu-id="68b26-203">数据源来自邮件的主题。</span><span class="sxs-lookup"><span data-stu-id="68b26-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="68b26-204">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-204">Requirements</span></span>

|<span data-ttu-id="68b26-205">要求</span><span class="sxs-lookup"><span data-stu-id="68b26-205">Requirement</span></span>| <span data-ttu-id="68b26-206">值</span><span class="sxs-lookup"><span data-stu-id="68b26-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="68b26-207">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="68b26-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="68b26-208">1.0</span><span class="sxs-lookup"><span data-stu-id="68b26-208">1.0</span></span>|
|[<span data-ttu-id="68b26-209">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="68b26-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="68b26-210">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="68b26-210">Compose or read</span></span>|